"""한컷영단어 로컬 AI 연동 서버. Python 3 표준 라이브러리만 사용.
사용: OPENAI_API_KEY 환경변수 설정 후 python server.py
이 데모는 loopback에만 바인딩되며 실제 운영용 인증·DB 서버를 대신하지 않습니다.
"""
import base64
import json
import os
from pathlib import Path
import re
import mimetypes
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
from urllib.parse import quote, urlparse, parse_qs
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from urllib.request import Request, urlopen
from urllib.error import HTTPError, URLError
from datetime import datetime, timezone, timedelta

ROOT = Path(__file__).resolve().parent
# Optional private settings at parent directory, not inside the served webroot.
# Existing environment variables take precedence. No keys are ever sent to browsers.
def load_private_env():
    path = ROOT.parent / 'onecut.private.env'
    if not path.is_file(): return
    for line in path.read_text(encoding='utf-8-sig').splitlines():
        entry=line.strip()
        if not entry or entry.startswith('#') or '=' not in entry: continue
        name,value=entry.split('=',1)
        name=name.strip(); value=value.strip()
        if name not in ('OPENAI_API_KEY','SUPABASE_URL','SUPABASE_ANON_KEY','SUPABASE_SERVICE_ROLE_KEY','ONECUT_TEXT_MODEL','ONECUT_IMAGE_MODEL','ONECUT_PORT'): continue
        if value.startswith(('\"',"'")) and len(value)>=2 and value[-1]==value[0]: value=value[1:-1]
        if value and name not in os.environ: os.environ[name]=value
load_private_env()
# 로컬은 127.0.0.1 기본, 클라우드(Render 등)는 PORT가 주어지므로 0.0.0.0에 바인딩
HOST = '0.0.0.0' if os.environ.get('PORT') else os.environ.get('HOST', '127.0.0.1')
PORT = int(os.environ.get('PORT') or os.environ.get('ONECUT_PORT', '8766'))
AI_MODEL = os.environ.get('ONECUT_TEXT_MODEL', 'gpt-4.1-mini')
IMAGE_MODEL = os.environ.get('ONECUT_IMAGE_MODEL', 'gpt-image-2.5-flare')
MAX_BODY = 75_000_000
SUPABASE_URL = os.environ.get('SUPABASE_URL', '').rstrip('/')
SUPABASE_PUBLIC_KEY = os.environ.get('SUPABASE_ANON_KEY', '')
SUPABASE_SERVICE_KEY = os.environ.get('SUPABASE_SERVICE_ROLE_KEY', '')
CLOUD_ENABLED = bool(SUPABASE_URL and SUPABASE_PUBLIC_KEY and SUPABASE_SERVICE_KEY)
VERSION = '5.3.0'
LOG_DIR = ROOT / 'logs'
LOG_PATH = LOG_DIR / 'onecut.log'

def log_event(kind, message, **extra):
    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        row={'at':datetime.now(timezone.utc).isoformat(),'kind':str(kind)[:40],'message':str(message)[:500]}
        if extra: row['extra']={str(k)[:60]:str(v)[:300] for k,v in extra.items()}
        with LOG_PATH.open('a',encoding='utf-8') as f: f.write(json.dumps(row,ensure_ascii=False)+'\n')
    except Exception: pass
if SUPABASE_URL and (urlparse(SUPABASE_URL).scheme != 'https' or not SUPABASE_URL.endswith('.supabase.co')):
    raise RuntimeError('SUPABASE_URL은 본인의 https://...supabase.co 주소여야 합니다.')

# v5.2.5 Excel preview parser (Python 표준 라이브러리만 사용)
def _xlsx_col_index(ref):
    letters=''.join(ch for ch in str(ref) if ch.isalpha()).upper()
    n=0
    for ch in letters:n=n*26+(ord(ch)-64)
    return max(0,n-1)

def _xlsx_text(node, shared):
    ns='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    typ=node.attrib.get('t','')
    if typ=='inlineStr':
        return ''.join(t.text or '' for t in node.findall('.//'+ns+'t'))
    v=node.find(ns+'v')
    if v is None:return ''
    raw=v.text or ''
    if typ=='s':
        try:return shared[int(raw)]
        except Exception:return ''
    return raw

def parse_content_xlsx_preview(body):
    raw=str(body.get('file_base64','') or '').strip()
    if not raw:raise ValueError('엑셀 파일 데이터가 없습니다.')
    if ',' in raw and raw.lower().startswith('data:'):raw=raw.split(',',1)[1]
    try:data=base64.b64decode(raw,validate=True)
    except Exception as exc:raise ValueError('엑셀 파일을 읽을 수 없습니다.') from exc
    if len(data)>6_000_000:raise ValueError('6MB 이하의 .xlsx 파일만 가능합니다.')
    try:
        z=zipfile.ZipFile(BytesIO(data))
        ns='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
        shared=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            root=ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall(ns+'si'):
                shared.append(''.join(t.text or '' for t in si.findall('.//'+ns+'t')))
        sheet_name='Sheet1'; sheet_path='xl/worksheets/sheet1.xml'
        if 'xl/workbook.xml' in z.namelist():
            wbroot=ET.fromstring(z.read('xl/workbook.xml'))
            sheets=wbroot.find(ns+'sheets')
            if sheets is not None and len(sheets):
                sheet_name=sheets[0].attrib.get('name','Sheet1')
        root=ET.fromstring(z.read(sheet_path))
        grid=[]
        for row in root.findall('.//'+ns+'row'):
            vals={}
            for c in row.findall(ns+'c'):
                vals[_xlsx_col_index(c.attrib.get('r','A1'))]=_xlsx_text(c,shared).strip()
            if vals:
                last=max(vals); grid.append([vals.get(i,'') for i in range(last+1)])
        if not grid:raise ValueError('엑셀에 데이터가 없습니다.')
    except ValueError:raise
    except Exception as exc:raise ValueError('엑셀 첫 시트를 열 수 없습니다.') from exc
    headers=[str(x or '').strip() for x in grid[0]]
    norm=lambda v:re.sub(r'\s+','',str(v or '')).lower()
    hmap={norm(h):i for i,h in enumerate(headers) if norm(h)}
    def pick(*names):
        for name in names:
            if norm(name) in hmap:return hmap[norm(name)]
        return None
    def cell(values,*names):
        i=pick(*names)
        return str(values[i] if i is not None and i<len(values) else '').strip()
    required=[('영어 단어',('영어 단어','단어','영단어','word')),('뜻',('뜻','의미','meaning')),('연상',('연상','발음연상','연상문장','mnemonic')),('이미지 장면',('이미지 장면','장면','이미지장면','scene'))]
    for label,names in required:
        if pick(*names) is None:raise ValueError('필수 열 누락: '+label)
    out=[];skipped=0
    for excel_row_num,values in enumerate(grid[1:],start=2):
        if not any(str(v or '').strip() for v in values):continue
        word=cell(values,'영어 단어','단어','영단어','word');meaning=cell(values,'뜻','의미','meaning')
        if not word or not meaning:skipped+=1;continue
        row={
          'word':word,'ipa':cell(values,'IPA 발음','IPA','발음','ipa'),
          'hangul':cell(values,'한글 발음연상','한글 발음 분해','한글 발음','hangul'),
          'stress_mark':cell(values,'강세 표시','강세표시','stress mark','stress_mark'),
          'meaning':meaning,'example':cell(values,'영어 예문','예문','example'),
          'translation':cell(values,'한글 번역','예문 번역','번역','translation'),
          'mnemonic':cell(values,'연상','발음연상','연상문장','mnemonic'),
          'speech':cell(values,'말풍선','대사','말풍선 대사','sceneSpeech'),
          'scene':cell(values,'이미지 장면','장면','이미지장면','scene'),
          'footer':cell(values,'하단 설명','설명','footer'),'status':cell(values,'검토 상태','상태','status'),
          'image_file':cell(values,'이미지 파일명','업로드 파일명','파일명','image file','image_file'),
          'excel_row':excel_row_num}
        missing=[k for k in ('word','meaning','mnemonic','scene','stress_mark') if not row[k]]
        mark=row['stress_mark']; hit=re.fullmatch(r'([^\[\]]*)\[([^\[\]]+)\]([^\[\]]*)',mark) if mark else None
        if mark and (not hit or (hit.group(1)+hit.group(2)+hit.group(3)).lower()!=word.lower()): missing.append('stress_mark_format')
        row['valid']=not missing;row['missing']=missing;out.append(row)
    invalid=[{'excel_row':x['excel_row'],'word':x['word'],'meaning':x['meaning'],'missing':x['missing']} for x in out if not x['valid']]
    return {'sheet':sheet_name,'headers':headers,'rows':out,'count':len(out),'total_rows':len(out),'invalid_rows':invalid,'skipped':skipped}


def supa_request(path, method='GET', data=None, headers=None, service=False, timeout=20, raw=False):
    """Backend-only Supabase proxy. Never give service-role key to the browser."""
    key=SUPABASE_SERVICE_KEY if service else SUPABASE_PUBLIC_KEY
    h={'apikey':key,'Authorization':'Bearer '+key,'Accept':'application/json'}
    if headers:h.update(headers)
    if data is not None and not raw:
        data=json.dumps(data,ensure_ascii=False).encode('utf-8');h['Content-Type']='application/json'
    req=Request(SUPABASE_URL+path,data=data,headers=h,method=method)
    try:
        with urlopen(req,timeout=timeout) as resp:
            body=resp.read()
            return json.loads(body) if body else {}
    except HTTPError as err:
        try:
            detail=json.loads(err.read().decode('utf-8'))
            message=detail.get('message') or detail.get('msg') or detail.get('error_description') or detail.get('error') or str(detail)
        except Exception:message=err.reason
        raise ValueError('Supabase '+str(err.code)+': '+str(message)[:220]) from err
    except URLError as err:
        raise ValueError('Supabase 연결 실패: '+str(err.reason)) from err




# Student APIs use the student's own access token and PostgreSQL RLS.
def student_identity(bearer):
    if not CLOUD_ENABLED: raise ValueError('학생 클라우드 서비스가 설정되지 않았습니다.')
    if not isinstance(bearer,str) or not bearer.startswith('Bearer '): raise PermissionError('학생 로그인이 필요합니다.')
    token=bearer[7:].strip()
    if not token or len(token)>4096:raise PermissionError('학생 로그인 토큰이 올바르지 않습니다.')
    user=supa_request('/auth/v1/user',headers={'Authorization':'Bearer '+token})
    uid=user.get('id','')
    if not uid:raise PermissionError('학생 로그인 정보가 유효하지 않습니다.')
    return token,{'id':uid,'email':user.get('email','')}

def student_request(token,method='GET',payload=None):
    path='/rest/v1/onecut_student_progress?select=progress%2Cupdated_at'
    if method=='GET':
        result=supa_request(path,headers={'Authorization':'Bearer '+token})
        return result[0].get('progress',{}) if result else {}
    if not isinstance(payload,dict) or len(payload)>3000:raise ValueError('학습기록 형식 또는 개수 오류')
    valid={}
    for k,p in payload.items():
        if not isinstance(k,str) or not re.fullmatch('[A-Za-z0-9_-]{1,90}',k):raise ValueError('단어 ID 오류')
        if not isinstance(p,dict):raise ValueError('학습기록 값 오류')
        count=p.get('attempts',0)
        if type(count)!=int or count<0 or count>100000:raise ValueError('학습 횟수 오류')
        value=p.get('value','again')
        if value not in ('again','hard','good','easy'):raise ValueError('암기 상태 오류')
        history=p.get('history',[])
        if not isinstance(history,list):history=[]
        safe_history=[]
        for event in history[-90:]:
            if not isinstance(event,dict):continue
            stamp=event.get('at','')
            rating=event.get('result','')
            if isinstance(stamp,str) and len(stamp)<=50 and rating in ('again','hard','good','easy'):
                safe_history.append({'at':stamp,'result':rating})
        success_streak=max(0,min(10000,int(p.get('successStreak',0) or 0)))
        lapses=max(0,min(10000,int(p.get('lapses',0) or 0)))
        try: interval_days=max(0.0,min(3650.0,float(p.get('intervalDays',0) or 0)))
        except (TypeError,ValueError): interval_days=0.0
        valid[k]={'attempts':count,'value':value,'favorite':p.get('favorite') is True,
            'updated':str(p.get('updated',''))[:50],'nextReview':str(p.get('nextReview',''))[:50],
            'successStreak':success_streak,'lapses':lapses,'intervalDays':interval_days,
            'mastered':p.get('mastered') is True,'history':safe_history}
    if len(json.dumps(valid,ensure_ascii=False))>500000:raise ValueError('학습기록이 너무 큽니다.')
    # user_id is derived from validated auth token, not trusted client input.
    user=supa_request('/auth/v1/user',headers={'Authorization':'Bearer '+token})
    return supa_request('/rest/v1/onecut_student_progress?on_conflict=user_id',method='POST',
        data={'user_id':user['id'],'progress':valid},headers={
        'Authorization':'Bearer '+token,'Prefer':'resolution=merge-duplicates,return=representation'})

def require_admin(bearer):
    if not CLOUD_ENABLED:raise ValueError('클라우드 모드가 아닙니다.')
    if not isinstance(bearer,str) or not bearer.startswith('Bearer '):raise PermissionError('관리자 로그인 필요')
    token=bearer[7:].strip()
    if not token or len(token)>4096:raise PermissionError('올바른 관리자 로그인 토큰이 아닙니다.')
    try:
        user=supa_request('/auth/v1/user',headers={'Authorization':'Bearer '+token})
        uid=user.get('id','')
        if not uid:raise PermissionError('관리자 로그인 실패')
        rows=supa_request('/rest/v1/app_admins?user_id=eq.'+quote(uid,safe='')+'&select=user_id',service=True)
        if not isinstance(rows,list) or not rows:raise PermissionError('해당 계정에는 관리자 권한이 없습니다.')
        return {'id':uid,'email':user.get('email','')}
    except ValueError as exc:raise PermissionError('관리자 인증 오류: '+str(exc)) from exc



def _parse_iso(value):
    if not isinstance(value,str) or not value:return None
    try:return datetime.fromisoformat(value.replace('Z','+00:00'))
    except ValueError:return None



def _assignment_student_state(progress, card_ids, now=None):
    """Sanitized live completion state for one assignment/student."""
    now=now or datetime.now(timezone.utc)
    ids=[str(x) for x in (card_ids or []) if isinstance(x,str)]
    total=len(ids); learned=mastered=weak=due=0
    for cid in ids:
        item=progress.get(cid) if isinstance(progress,dict) else None
        if not isinstance(item,dict): continue
        attempts=item.get('attempts',0)
        if isinstance(attempts,int) and attempts>0: learned+=1
        value=str(item.get('value') or '')
        try: streak=max(0,int(item.get('successStreak',0) or 0))
        except Exception: streak=0
        try: interval=max(0,float(item.get('intervalDays',0) or 0))
        except Exception: interval=0
        is_mastered=item.get('mastered') is True or (streak>=4 and interval>=30 and value in ('good','easy'))
        if is_mastered: mastered+=1
        if attempts and value in ('again','hard'): weak+=1
        nr=_parse_iso(item.get('nextReview'))
        if nr:
            if nr.tzinfo is None: nr=nr.replace(tzinfo=timezone.utc)
            if nr<=now: due+=1
    completion=round(learned/total*100) if total else 0
    mastery=round(mastered/total*100) if total else 0
    if learned<=0: status='not_started'
    elif learned<total: status='in_progress'
    elif weak>0 or due>0: status='attention'
    else: status='completed'
    return {'total':total,'learned':learned,'completion':completion,'mastered':mastered,'mastery':mastery,
        'weak':weak,'due':due,'status':status}


def admin_groups_overview():
    """Return groups, assignments and live assignment completion. Admin only."""
    groups=supa_request('/rest/v1/onecut_groups?select=id,name,created_at&order=created_at.asc',service=True)
    members=supa_request('/rest/v1/onecut_group_members?select=group_id,user_id,created_at&order=created_at.asc',service=True)
    assignments=supa_request('/rest/v1/onecut_assignments?select=id,title,card_ids,target_type,target_id,assigned_at,due_at,active&order=assigned_at.desc',service=True)
    progress_rows=supa_request('/rest/v1/onecut_student_progress?select=user_id,progress,updated_at',service=True)
    if not isinstance(groups,list): groups=[]
    if not isinstance(members,list): members=[]
    if not isinstance(assignments,list): assignments=[]
    if not isinstance(progress_rows,list): progress_rows=[]
    emails={}
    try:
        payload=supa_request('/auth/v1/admin/users?page=1&per_page=1000',service=True)
        for u in (payload.get('users',[]) if isinstance(payload,dict) else []):
            if isinstance(u,dict) and u.get('id'): emails[str(u['id'])]=str(u.get('email') or '')[:254]
    except Exception: pass
    by_group={}
    for m in members:
        gid=str(m.get('group_id') or ''); uid=str(m.get('user_id') or '')
        by_group.setdefault(gid,[]).append({'user_id':uid,'email':emails.get(uid,'')})
    progress_by_user={str(r.get('user_id') or ''):(r.get('progress') if isinstance(r.get('progress'),dict) else {}) for r in progress_rows if isinstance(r,dict)}
    out=[]
    for g in groups:
        gid=str(g.get('id') or '')
        out.append({'id':gid,'name':str(g.get('name') or '')[:80],'created_at':g.get('created_at'),'members':by_group.get(gid,[])})
    assignment_out=[]; now=datetime.now(timezone.utc)
    for raw in assignments:
        if not isinstance(raw,dict): continue
        a=dict(raw); ids=a.get('card_ids') if isinstance(a.get('card_ids'),list) else []
        target_type=str(a.get('target_type') or ''); target_id=str(a.get('target_id') or '')
        target_users=[target_id] if target_type=='student' and target_id else [m['user_id'] for m in by_group.get(target_id,[])]
        states=[]
        for uid in target_users:
            state=_assignment_student_state(progress_by_user.get(uid,{}),ids,now)
            state.update({'user_id':uid,'email':emails.get(uid,'')})
            states.append(state)
        counts={'students':len(states),'not_started':0,'in_progress':0,'attention':0,'completed':0}
        for st in states: counts[st['status']]=counts.get(st['status'],0)+1
        avg=round(sum(st['completion'] for st in states)/len(states)) if states else 0
        mastery=round(sum(st['mastery'] for st in states)/len(states)) if states else 0
        a['progress']={'counts':counts,'average_completion':avg,'average_mastery':mastery,'students':states}
        assignment_out.append(a)
    return {'groups':out,'assignments':assignment_out}

def admin_groups_action(actor,body):
    action=str(body.get('action') or '')
    if action=='create_group':
        name=str(body.get('name') or '').strip()
        if not name or len(name)>80: raise ValueError('그룹 이름을 1~80자로 입력하세요.')
        return supa_request('/rest/v1/onecut_groups',method='POST',data={'name':name,'created_by':actor['id']},service=True,headers={'Prefer':'return=representation'})
    if action in ('add_member','remove_member'):
        gid=str(body.get('group_id') or ''); uid=str(body.get('user_id') or '')
        if not re.fullmatch(r'[0-9a-fA-F-]{36}',gid) or not re.fullmatch(r'[0-9a-fA-F-]{36}',uid): raise ValueError('그룹 또는 학생 ID 오류')
        path='/rest/v1/onecut_group_members?group_id=eq.'+quote(gid,safe='')+'&user_id=eq.'+quote(uid,safe='')
        if action=='remove_member': return supa_request(path,method='DELETE',service=True,headers={'Prefer':'return=representation'})
        return supa_request('/rest/v1/onecut_group_members',method='POST',data={'group_id':gid,'user_id':uid},service=True,headers={'Prefer':'resolution=merge-duplicates,return=representation'})
    if action=='create_assignment':
        title=str(body.get('title') or '').strip()
        target_type=str(body.get('target_type') or '')
        target_id=str(body.get('target_id') or '')
        ids=body.get('card_ids')
        due=str(body.get('due_at') or '').strip()
        if not title or len(title)>100: raise ValueError('학습세트 제목을 입력하세요.')
        if target_type not in ('student','group'): raise ValueError('배정 대상을 확인하세요.')
        if not re.fullmatch(r'[0-9a-fA-F-]{36}',target_id): raise ValueError('배정 대상 ID 오류')
        if not isinstance(ids,list) or not 1<=len(ids)<=50 or any(not isinstance(x,str) or not re.fullmatch(r'[A-Za-z0-9_-]{1,90}',x) for x in ids): raise ValueError('학습세트 단어는 1~50개여야 합니다.')
        payload={'title':title,'card_ids':ids,'target_type':target_type,'target_id':target_id,'created_by':actor['id'],'active':True}
        if due: payload['due_at']=due[:50]
        return supa_request('/rest/v1/onecut_assignments',method='POST',data=payload,service=True,headers={'Prefer':'return=representation'})
    if action=='deactivate_assignment':
        aid=str(body.get('assignment_id') or '')
        if not re.fullmatch(r'[0-9a-fA-F-]{36}',aid): raise ValueError('배정 ID 오류')
        return supa_request('/rest/v1/onecut_assignments?id=eq.'+quote(aid,safe=''),method='PATCH',data={'active':False},service=True,headers={'Prefer':'return=representation'})
    raise ValueError('지원하지 않는 그룹 작업입니다.')

def student_assignments(user_id):
    uid=str(user_id or '')
    direct=supa_request('/rest/v1/onecut_assignments?active=eq.true&target_type=eq.student&target_id=eq.'+quote(uid,safe='')+'&select=id,title,card_ids,assigned_at,due_at&order=assigned_at.desc',service=True)
    memberships=supa_request('/rest/v1/onecut_group_members?user_id=eq.'+quote(uid,safe='')+'&select=group_id',service=True)
    gids=[str(x.get('group_id')) for x in memberships] if isinstance(memberships,list) else []
    group_rows=[]
    for gid in gids[:50]:
        rows=supa_request('/rest/v1/onecut_assignments?active=eq.true&target_type=eq.group&target_id=eq.'+quote(gid,safe='')+'&select=id,title,card_ids,assigned_at,due_at&order=assigned_at.desc',service=True)
        if isinstance(rows,list): group_rows.extend(rows)
    items=(direct if isinstance(direct,list) else [])+group_rows
    seen=set(); out=[]
    for a in sorted(items,key=lambda x:str(x.get('assigned_at') or ''),reverse=True):
        aid=str(a.get('id') or '')
        if aid and aid not in seen: seen.add(aid); out.append(a)
    return {'assignments':out[:100]}


def admin_student_overview():
    """Aggregate student progress with the service role. Raw progress is never returned."""
    if not CLOUD_ENABLED:raise ValueError('학생 클라우드 서비스가 설정되지 않았습니다.')
    rows=supa_request('/rest/v1/onecut_student_progress?select=user_id,progress,updated_at&order=updated_at.desc',service=True)
    if not isinstance(rows,list):rows=[]
    emails={}
    # Supabase Auth Admin endpoint can list users with a service-role key.
    try:
        payload=supa_request('/auth/v1/admin/users?page=1&per_page=1000',service=True)
        users=payload.get('users',[]) if isinstance(payload,dict) else []
        for u in users:
            if isinstance(u,dict) and u.get('id'):
                emails[str(u['id'])]=str(u.get('email') or '')[:254]
    except Exception:
        # The learning overview remains useful even if auth metadata is temporarily unavailable.
        pass
    try:
        published=cloud_published()
        pub_cards=published.get('cards',[]) if isinstance(published,dict) else published
        total_words=len(pub_cards) if isinstance(pub_cards,list) else 0
    except Exception:total_words=0
    now=datetime.now(timezone.utc);kst=timezone(timedelta(hours=9));today=now.astimezone(kst).date()
    students=[]
    for row in rows:
        if not isinstance(row,dict):continue
        uid=str(row.get('user_id') or '')
        progress=row.get('progress') if isinstance(row.get('progress'),dict) else {}
        learned=mastered=weak=due=today_count=events=correct=0
        last_active=_parse_iso(row.get('updated_at'))
        next_due=None
        for item in progress.values():
            if not isinstance(item,dict):continue
            attempts=item.get('attempts',0)
            if isinstance(attempts,int) and attempts>0:learned+=1
            value=item.get('value')
            streak=max(0,int(item.get('successStreak',0) or 0)) if str(item.get('successStreak',0)).lstrip('-').isdigit() else 0
            try:interval=float(item.get('intervalDays',0) or 0)
            except (TypeError,ValueError):interval=0
            is_mastered=item.get('mastered') is True or (streak>=4 and interval>=30 and value in ('good','easy'))
            if is_mastered:mastered+=1
            if value in ('again','hard'):weak+=1
            nr=_parse_iso(item.get('nextReview'))
            if nr:
                if nr.tzinfo is None:nr=nr.replace(tzinfo=timezone.utc)
                if nr<=now:due+=1
                elif next_due is None or nr<next_due:next_due=nr
            updated=_parse_iso(item.get('updated'))
            if updated:
                if updated.tzinfo is None:updated=updated.replace(tzinfo=timezone.utc)
                if last_active is None or updated>last_active:last_active=updated
                if updated.astimezone(kst).date()==today:today_count+=1
            hist=item.get('history',[])
            if isinstance(hist,list):
                for h in hist[-120:]:
                    if not isinstance(h,dict):continue
                    result=h.get('result')
                    if result in ('again','hard','good','easy'):
                        events+=1
                        if result in ('good','easy'):correct+=1
        accuracy=round(correct/events*100) if events else None
        mastery_rate=round(mastered/learned*100) if learned else 0
        label=emails.get(uid) or ('학생 '+uid[:8] if uid else '학생')
        students.append({'user_id':uid,'email':emails.get(uid,''),'label':label,'learned':learned,
            'total_words':total_words,'mastered':mastered,'mastery_rate':mastery_rate,'weak':weak,
            'due':due,'today':today_count,'accuracy':accuracy,'events':events,
            'last_active':last_active.isoformat() if last_active else '',
            'next_due':next_due.isoformat() if next_due else ''})
    students.sort(key=lambda x:(x['weak']>0 or x['due']>0, x['weak']+x['due'], x['last_active']), reverse=True)
    summary={'students':len(students),'active_today':sum(1 for x in students if x['today']>0),
        'need_attention':sum(1 for x in students if x['weak']>0 or x['due']>0),
        'mastered_total':sum(x['mastered'] for x in students),'published_words':total_words}
    return {'summary':summary,'students':students}


def admin_student_detail(user_id):
    """Administrator-only, sanitized learning detail for one student."""
    if not CLOUD_ENABLED:raise ValueError('학생 클라우드 서비스가 설정되지 않았습니다.')
    uid=str(user_id or '').strip()
    if not re.fullmatch(r'[0-9a-fA-F-]{8,64}',uid):raise ValueError('학생 ID 형식이 올바르지 않습니다.')
    rows=supa_request('/rest/v1/onecut_student_progress?user_id=eq.'+quote(uid,safe='')+'&select=user_id,progress,updated_at&limit=1',service=True)
    if not isinstance(rows,list) or not rows:raise ValueError('학생 학습기록을 찾을 수 없습니다.')
    row=rows[0]; progress=row.get('progress') if isinstance(row.get('progress'),dict) else {}
    email=''
    try:
        u=supa_request('/auth/v1/admin/users/'+quote(uid,safe=''),service=True)
        if isinstance(u,dict):email=str(u.get('email') or '')[:254]
    except Exception:pass
    try:
        pub=cloud_published(); card_list=pub.get('cards',[]) if isinstance(pub,dict) else (pub if isinstance(pub,list) else [])
    except Exception:card_list=[]
    cards={str(c.get('id')):c for c in card_list if isinstance(c,dict) and c.get('id')}
    now=datetime.now(timezone.utc); kst=timezone(timedelta(hours=9)); today=now.astimezone(kst).date()
    details=[]; daily={}
    def add_day(dt,result=None):
        day=dt.astimezone(kst).date().isoformat(); d=daily.setdefault(day,{'events':0,'correct':0,'words':set()}); d['events']+=1
        if result in ('good','easy'):d['correct']+=1
    for cid,item in progress.items():
        if not isinstance(item,dict):continue
        c=cards.get(str(cid),{}); attempts=max(0,int(item.get('attempts',0) or 0)) if str(item.get('attempts',0)).isdigit() else 0
        value=str(item.get('value') or '')
        try:streak=max(0,int(item.get('successStreak',0) or 0))
        except Exception:streak=0
        try:lapses=max(0,int(item.get('lapses',0) or 0))
        except Exception:lapses=0
        try:interval=max(0,float(item.get('intervalDays',0) or 0))
        except Exception:interval=0
        mastered=item.get('mastered') is True or (streak>=4 and interval>=30 and value in ('good','easy'))
        nr=_parse_iso(item.get('nextReview')); updated=_parse_iso(item.get('updated'))
        if nr and nr.tzinfo is None:nr=nr.replace(tzinfo=timezone.utc)
        if updated and updated.tzinfo is None:updated=updated.replace(tzinfo=timezone.utc)
        due=bool(nr and nr<=now)
        hist=item.get('history',[]) if isinstance(item.get('history'),list) else []
        wrong_recent=0; last_result=value
        for h in hist[-120:]:
            if not isinstance(h,dict):continue
            dt=_parse_iso(h.get('at')); result=h.get('result')
            if dt:
                if dt.tzinfo is None:dt=dt.replace(tzinfo=timezone.utc)
                add_day(dt,result)
            if result in ('again','hard'):wrong_recent+=1
            if result in ('again','hard','good','easy'):last_result=result
        weakness=(5 if value=='again' else 3 if value=='hard' else 0)+(2 if due else 0)+min(6,lapses*2)+min(4,wrong_recent)
        status='mastered' if mastered else 'weak' if value in ('again','hard') else 'due' if due else 'learning' if attempts else 'new'
        details.append({'id':str(cid),'word':str(c.get('word') or cid),'pos':str(c.get('pos') or ''),'meaning':str(c.get('meaning') or ''),
            'attempts':attempts,'value':value,'last_result':last_result,'success_streak':streak,'lapses':lapses,'interval_days':round(interval,2),
            'mastered':mastered,'due':due,'status':status,'weakness':weakness,'next_review':nr.isoformat() if nr else '',
            'updated':updated.isoformat() if updated else ''})
    details.sort(key=lambda x:(x['weakness'],x['due'],x['updated']),reverse=True)
    weak_top=[x for x in details if x['weakness']>0 and not x['mastered']][:10]
    recent_wrong=[x for x in details if x['last_result'] in ('again','hard')][:10]
    days=[]
    for offset in range(29,-1,-1):
        day=(today-timedelta(days=offset)).isoformat(); d=daily.get(day,{'events':0,'correct':0}); ev=d.get('events',0); co=d.get('correct',0)
        days.append({'date':day,'events':ev,'correct':co,'accuracy':round(co/ev*100) if ev else None})
    def period(n):
        arr=days[-n:]; ev=sum(x['events'] for x in arr); co=sum(x['correct'] for x in arr)
        return {'events':ev,'correct':co,'accuracy':round(co/ev*100) if ev else None,'active_days':sum(1 for x in arr if x['events']>0)}
    learned=sum(1 for x in details if x['attempts']>0); mastered=sum(1 for x in details if x['mastered']); weak=sum(1 for x in details if x['status']=='weak'); due=sum(1 for x in details if x['due'])
    return {'student':{'user_id':uid,'email':email,'label':email or ('학생 '+uid[:8]),'updated_at':row.get('updated_at','')},
        'summary':{'learned':learned,'mastered':mastered,'mastery_rate':round(mastered/learned*100) if learned else 0,'weak':weak,'due':due,'total_words':len(cards)},
        'periods':{'days7':period(7),'days30':period(30)},'daily':days,'weak_top':weak_top,'recent_wrong':recent_wrong,'words':details}




def _safe_filter(value, allowed=None, limit=80):
    value=str(value or '').strip()[:limit]
    if allowed is not None and value and value not in allowed: raise ValueError('분류 조건이 올바르지 않습니다.')
    return value

def _published_cards_list():
    pub=cloud_published(); return pub.get('cards',[]) if isinstance(pub,dict) else (pub if isinstance(pub,list) else [])

def _card_meta_rows():
    try:
        rows=supa_request('/rest/v1/onecut_card_meta?select=card_id,difficulty,school_level,exam_scope,topic,tags,updated_at',service=True)
        return rows if isinstance(rows,list) else []
    except Exception:
        return []

def _card_meta_map():
    return {str(r.get('card_id')):r for r in _card_meta_rows() if isinstance(r,dict) and r.get('card_id')}

def admin_catalog():
    cards=_published_cards_list(); mm=_card_meta_map(); out=[]
    for c in cards:
        if not isinstance(c,dict): continue
        x={k:c.get(k) for k in ('id','word','pos','meaning','status')}; m=mm.get(str(c.get('id')), {})
        x['meta']={'difficulty':m.get('difficulty') or '미분류','school_level':m.get('school_level') or '미분류','exam_scope':m.get('exam_scope') or '', 'topic':m.get('topic') or '', 'tags':m.get('tags') if isinstance(m.get('tags'),list) else []}
        out.append(x)
    return {'cards':out,'counts':{'published':len(out),'classified':sum(1 for x in out if x['meta']['difficulty']!='미분류')}}

def admin_catalog_action(actor, body):
    action=str(body.get('action') or '')
    if action=='save_meta':
        cid=str(body.get('card_id') or '').strip()
        if not re.fullmatch(r'[A-Za-z0-9_-]{1,90}',cid): raise ValueError('단어 ID 오류')
        difficulty=_safe_filter(body.get('difficulty'), {'미분류','초급','중급','고급'}) or '미분류'
        school=_safe_filter(body.get('school_level'), {'미분류','초등','중등','고등','수능'}) or '미분류'
        exam=_safe_filter(body.get('exam_scope'), None, 100); topic=_safe_filter(body.get('topic'), None, 100)
        tags=body.get('tags',[]); tags=[str(x).strip()[:30] for x in tags if str(x).strip()][:12] if isinstance(tags,list) else []
        payload={'card_id':cid,'difficulty':difficulty,'school_level':school,'exam_scope':exam,'topic':topic,'tags':tags,'updated_by':actor['id'],'updated_at':datetime.now(timezone.utc).isoformat()}
        return supa_request('/rest/v1/onecut_card_meta?on_conflict=card_id',method='POST',data=payload,service=True,headers={'Prefer':'resolution=merge-duplicates,return=representation'})
    raise ValueError('지원하지 않는 분류 작업입니다.')

def _progress_for_user(uid):
    rows=supa_request('/rest/v1/onecut_student_progress?user_id=eq.'+quote(uid,safe='')+'&select=progress&limit=1',service=True)
    return (rows[0].get('progress') if isinstance(rows,list) and rows and isinstance(rows[0].get('progress'),dict) else {})

def _existing_active_card_ids(uid):
    try:
        out=set()
        for a in student_assignments(uid).get('assignments',[]):
            for cid in (a.get('card_ids') or []): out.add(str(cid))
        return out
    except Exception:return set()

def _filtered_cards(filters=None):
    filters=filters or {}; mm=_card_meta_map(); out=[]
    for c in _published_cards_list():
        if not isinstance(c,dict) or not c.get('id'): continue
        m=mm.get(str(c['id']),{}); ok=True
        for key in ('difficulty','school_level'):
            want=str(filters.get(key) or '').strip()
            if want and want!='전체' and str(m.get(key) or '미분류')!=want: ok=False
        want=str(filters.get('exam_scope') or '').strip().lower()
        if want and want not in str(m.get('exam_scope') or '').lower(): ok=False
        want=str(filters.get('topic') or '').strip().lower()
        if want and want not in str(m.get('topic') or '').lower(): ok=False
        if ok:
            x=dict(c); x['meta']=m; out.append(x)
    return out

def _recommend_filtered(uid, count, filters=None, exclude=None):
    count=max(1,min(20,int(count))); exclude=set(exclude or [])|_existing_active_card_ids(uid)
    progress=_progress_for_user(uid); pool=_filtered_cards(filters); now=datetime.now(timezone.utc)
    weak=[]; unseen=[]; fallback=[]
    for c in pool:
        cid=str(c.get('id')); p=progress.get(cid) if isinstance(progress,dict) else None
        if cid in exclude: continue
        if not isinstance(p,dict) or not p.get('attempts'):
            unseen.append(c); continue
        value=str(p.get('value') or ''); mastered=p.get('mastered') is True
        if mastered: continue
        try:lapses=max(0,int(p.get('lapses',0) or 0))
        except Exception:lapses=0
        nr=_parse_iso(p.get('nextReview')); due=bool(nr and (nr if nr.tzinfo else nr.replace(tzinfo=timezone.utc))<=now)
        score=(8 if value=='again' else 5 if value=='hard' else 0)+min(8,lapses*2)+(3 if due else 0)
        entry=(score,c,'모름·헷갈림' if value in ('again','hard') else '복습 시점 도래' if due else '암기 진행')
        if score>0: weak.append(entry)
        else: fallback.append(c)
    weak.sort(key=lambda x:x[0],reverse=True)
    weak_quota=min(len(weak), max(1,round(count*0.4)))
    chosen=[{'card':c,'kind':'취약','reason':reason,'score':score} for score,c,reason in weak[:weak_quota]]
    for c in unseen:
        if len(chosen)>=count: break
        chosen.append({'card':c,'kind':'신규','reason':'아직 학습하지 않은 단어','score':0})
    for score,c,reason in weak[weak_quota:]:
        if len(chosen)>=count: break
        chosen.append({'card':c,'kind':'취약','reason':reason,'score':score})
    for c in fallback:
        if len(chosen)>=count:break
        chosen.append({'card':c,'kind':'복습','reason':'암기 유지 복습','score':0})
    return chosen

def admin_curriculum_preview(body):
    gid=str(body.get('group_id') or '')
    if not re.fullmatch(r'[0-9a-fA-F-]{36}',gid): raise ValueError('그룹을 선택하세요.')
    count=max(1,min(20,int(body.get('count',5) or 5))); sessions=max(1,min(4,int(body.get('sessions',1) or 1)))
    filters={'difficulty':str(body.get('difficulty') or ''),'school_level':str(body.get('school_level') or ''),'exam_scope':str(body.get('exam_scope') or ''),'topic':str(body.get('topic') or '')}
    members=supa_request('/rest/v1/onecut_group_members?group_id=eq.'+quote(gid,safe='')+'&select=user_id',service=True)
    emails={}
    try:
        au=supa_request('/auth/v1/admin/users?page=1&per_page=1000',service=True)
        for u in (au.get('users',[]) if isinstance(au,dict) else []): emails[str(u.get('id'))]=str(u.get('email') or '')
    except Exception:pass
    plans=[]
    for m in (members if isinstance(members,list) else []):
        uid=str(m.get('user_id') or ''); exclude=set(); ss=[]
        for sn in range(sessions):
            rec=_recommend_filtered(uid,count,filters,exclude)
            ids=[str(x['card'].get('id')) for x in rec]; exclude.update(ids)
            ss.append({'session':sn+1,'card_ids':ids,'items':[{'id':str(x['card'].get('id')),'word':str(x['card'].get('word') or ''),'meaning':str(x['card'].get('meaning') or ''),'kind':x['kind'],'reason':x['reason']} for x in rec]})
        plans.append({'user_id':uid,'email':emails.get(uid,''),'sessions':ss})
    return {'group_id':gid,'count':count,'session_count':sessions,'filters':filters,'students':plans,'total_students':len(plans)}

def admin_curriculum_apply(actor,body):
    preview=admin_curriculum_preview(body); title=str(body.get('title') or '자동 커리큘럼').strip()[:80] or '자동 커리큘럼'; due_days=max(0,min(365,int(body.get('due_days',7) or 7)))
    created=[]
    for st in preview['students']:
        for sess in st['sessions']:
            if not sess['card_ids']: continue
            due=(datetime.now(timezone.utc)+timedelta(days=due_days*sess['session'])).isoformat() if due_days else ''
            payload={'title':f"{title} {sess['session']}회차",'card_ids':sess['card_ids'],'target_type':'student','target_id':st['user_id'],'created_by':actor['id'],'active':True}
            if due:payload['due_at']=due
            rows=supa_request('/rest/v1/onecut_assignments',method='POST',data=payload,service=True,headers={'Prefer':'return=representation'})
            created.extend(rows if isinstance(rows,list) else [])
    try:
        supa_request('/rest/v1/onecut_curricula',method='POST',data={'name':title,'group_id':preview['group_id'],'settings':{'count':preview['count'],'sessions':preview['session_count'],'filters':preview['filters'],'due_days':due_days},'created_by':actor['id'],'active':True},service=True,headers={'Prefer':'return=representation'})
    except Exception:pass
    return {'created_assignments':len(created),'students':preview['total_students'],'preview':preview}

def record_exam_attempt(uid, body):
    mode=_safe_filter(body.get('mode'), {'meaning','spelling','listening','mixed'}) or 'mixed'
    total=max(1,min(100,int(body.get('total',0) or 0))); correct=max(0,min(total,int(body.get('correct',0) or 0)))
    ids=body.get('card_ids',[]); ids=[str(x)[:90] for x in ids if isinstance(x,str)][:100] if isinstance(ids,list) else []
    answers=body.get('answers',[]); answers=answers[:100] if isinstance(answers,list) else []
    safe_answers=[]
    for x in answers:
        if isinstance(x,dict):safe_answers.append({k:str(v)[:120] for k,v in x.items() if k in ('card_id','type','result')})
    payload={'user_id':uid,'mode':mode,'total':total,'correct':correct,'score':round(correct/total*100),'card_ids':ids,'answers':safe_answers}
    try:return supa_request('/rest/v1/onecut_exam_attempts',method='POST',data=payload,service=True,headers={'Prefer':'return=representation'})
    except Exception as exc:raise ValueError('시험 기록 테이블 준비 필요: database/08_v5_platform.sql을 실행하세요. '+str(exc)) from exc

def admin_report(target_type,target_id,days=7):
    days=max(1,min(90,int(days))); target_type=str(target_type or '')
    uids=[]; label=''
    if target_type=='student': uids=[str(target_id)]; label='학생'
    elif target_type=='group':
        gid=str(target_id); rows=supa_request('/rest/v1/onecut_group_members?group_id=eq.'+quote(gid,safe='')+'&select=user_id',service=True); uids=[str(x.get('user_id')) for x in rows] if isinstance(rows,list) else []
        gs=supa_request('/rest/v1/onecut_groups?id=eq.'+quote(gid,safe='')+'&select=name&limit=1',service=True); label=str(gs[0].get('name') if isinstance(gs,list) and gs else '그룹')
    else: raise ValueError('리포트 대상을 확인하세요.')
    cutoff=datetime.now(timezone.utc)-timedelta(days=days); students=[]
    for uid in uids[:200]:
        try:d=admin_student_detail(uid)
        except Exception:continue
        summary=d.get('summary',{}); recent=d.get('daily',[])[-days:]
        exams=[]
        try:
            rows=supa_request('/rest/v1/onecut_exam_attempts?user_id=eq.'+quote(uid,safe='')+'&created_at=gte.'+quote(cutoff.isoformat(),safe='')+'&select=mode,total,correct,score,created_at&order=created_at.desc&limit=50',service=True)
            exams=rows if isinstance(rows,list) else []
        except Exception:pass
        avg_exam=round(sum(float(x.get('score',0) or 0) for x in exams)/len(exams)) if exams else None
        period=(d.get('periods',{}).get('days7') if days<=7 else d.get('periods',{}).get('days30')) or {}; summary=dict(summary); summary['accuracy']=period.get('accuracy'); summary['active_days']=period.get('active_days',0); students.append({'user_id':uid,'email':d.get('student',{}).get('email',''),'summary':summary,'trend':recent,'exam_count':len(exams),'exam_avg':avg_exam,'weak_top':d.get('weak_top',[])[:5]})
    n=len(students); aggregate={'students':n,'learned':sum(x['summary'].get('learned',0) for x in students),'mastered':sum(x['summary'].get('mastered',0) for x in students),'weak':sum(x['summary'].get('weak',0) for x in students),'due':sum(x['summary'].get('due',0) for x in students)}
    rates=[x['summary'].get('accuracy') for x in students if x['summary'].get('accuracy') is not None]; aggregate['accuracy']=round(sum(rates)/len(rates)) if rates else None
    er=[x['exam_avg'] for x in students if x['exam_avg'] is not None]; aggregate['exam_avg']=round(sum(er)/len(er)) if er else None
    return {'target_type':target_type,'target_id':target_id,'label':label,'days':days,'generated_at':datetime.now(timezone.utc).isoformat(),'aggregate':aggregate,'students':students}

def admin_backup():
    tables={'cards':'vocab_cards','images':'vocab_images','meta':'onecut_card_meta','groups':'onecut_groups','members':'onecut_group_members','assignments':'onecut_assignments','curricula':'onecut_curricula','student_progress':'onecut_student_progress','exam_attempts':'onecut_exam_attempts'}
    data={}
    for key,table in tables.items():
        try:
            rows=supa_request('/rest/v1/'+table+'?select=*',service=True); data[key]=rows if isinstance(rows,list) else []
        except Exception:data[key]=[]
    return {'version':VERSION,'exported_at':datetime.now(timezone.utc).isoformat(),'data':data}
def admin_student_recommendation(user_id, count=5):
    """Build a supervised next-learning recommendation: weak review first, then unseen published words."""
    if not CLOUD_ENABLED: raise ValueError('학생 클라우드 서비스가 설정되지 않았습니다.')
    uid=str(user_id or '').strip()
    if not re.fullmatch(r'[0-9a-fA-F-]{8,64}',uid): raise ValueError('학생 ID 형식이 올바르지 않습니다.')
    try: count=int(count)
    except Exception: count=5
    count=10 if count>=10 else 5
    rows=supa_request('/rest/v1/onecut_student_progress?user_id=eq.'+quote(uid,safe='')+'&select=progress&limit=1',service=True)
    progress=(rows[0].get('progress') if isinstance(rows,list) and rows and isinstance(rows[0].get('progress'),dict) else {})
    pub=cloud_published(); card_list=pub.get('cards',[]) if isinstance(pub,dict) else (pub if isinstance(pub,list) else [])
    cards=[c for c in card_list if isinstance(c,dict) and c.get('id')]
    by_id={str(c.get('id')):c for c in cards}
    now=datetime.now(timezone.utc)
    weak=[]
    for cid,item in progress.items():
        if str(cid) not in by_id or not isinstance(item,dict): continue
        attempts=item.get('attempts',0)
        try: attempts=max(0,int(attempts or 0))
        except Exception: attempts=0
        if attempts<=0: continue
        value=str(item.get('value') or '')
        try: streak=max(0,int(item.get('successStreak',0) or 0))
        except Exception: streak=0
        try: lapses=max(0,int(item.get('lapses',0) or 0))
        except Exception: lapses=0
        try: interval=max(0,float(item.get('intervalDays',0) or 0))
        except Exception: interval=0
        mastered=item.get('mastered') is True or (streak>=4 and interval>=30 and value in ('good','easy'))
        if mastered: continue
        nr=_parse_iso(item.get('nextReview'))
        if nr and nr.tzinfo is None: nr=nr.replace(tzinfo=timezone.utc)
        due=bool(nr and nr<=now)
        hist=item.get('history',[]) if isinstance(item.get('history'),list) else []
        recent_wrong=sum(1 for h in hist[-12:] if isinstance(h,dict) and h.get('result') in ('again','hard'))
        score=(7 if value=='again' else 5 if value=='hard' else 0)+(3 if due else 0)+min(8,lapses*2)+min(5,recent_wrong)
        if score<=0: continue
        c=by_id[str(cid)]
        reasons=[]
        if value=='again': reasons.append('최근 모름')
        elif value=='hard': reasons.append('최근 헷갈림')
        if due: reasons.append('복습 시점 도래')
        if lapses: reasons.append(f'망각 {lapses}회')
        if recent_wrong>=2: reasons.append('최근 반복 오답')
        weak.append({'id':str(cid),'word':str(c.get('word') or ''),'meaning':str(c.get('meaning') or ''),'score':score,'reason':' · '.join(reasons) or '취약 복습'})
    weak.sort(key=lambda x:(x['score'],x['word']),reverse=True)
    assigned_ids=set()
    try:
        for a in student_assignments(uid).get('assignments',[]):
            for cid in (a.get('card_ids') if isinstance(a.get('card_ids'),list) else []): assigned_ids.add(str(cid))
    except Exception: pass
    seen_ids=set()
    for cid,item in progress.items():
        if not isinstance(item,dict): continue
        try: attempts=int(item.get('attempts',0) or 0)
        except Exception: attempts=0
        if attempts>0: seen_ids.add(str(cid))
    new=[]
    for c in cards:
        cid=str(c.get('id') or '')
        if not cid or cid in seen_ids or cid in assigned_ids: continue
        new.append({'id':cid,'word':str(c.get('word') or ''),'meaning':str(c.get('meaning') or ''),'reason':'아직 학습하지 않은 공개 단어'})
    weak_limit=min(len(weak),3 if count==5 else 4)
    chosen_weak=weak[:weak_limit]
    chosen_ids={x['id'] for x in chosen_weak}
    need=count-len(chosen_weak)
    chosen_new=[x for x in new if x['id'] not in chosen_ids][:need]
    # If unseen words are insufficient, fill the remaining slots with additional weak words.
    if len(chosen_weak)+len(chosen_new)<count:
        extra=[x for x in weak[weak_limit:] if x['id'] not in chosen_ids][:count-len(chosen_weak)-len(chosen_new)]
        chosen_weak+=extra
    items=chosen_weak+chosen_new
    return {'student_id':uid,'count':len(items),'requested':count,'weak':chosen_weak,'new':chosen_new,
        'card_ids':[x['id'] for x in items],
        'summary':{'weak':len(chosen_weak),'new':len(chosen_new),'available_weak':len(weak),'available_new':len(new)},
        'policy':'완전 암기 단어는 제외하고 취약단어를 우선한 뒤 미학습 공개단어로 채웁니다. 관리자 승인 전에는 배정하지 않습니다.'}

def validate_cards(cards, previous_cards=None):
    old_by_id={c.get('id'):c for c in (previous_cards or []) if isinstance(c,dict)}
    if not isinstance(cards,list) or len(cards)>3000:raise ValueError('카드 목록 형식 또는 개수가 잘못되었습니다.')
    ids=set();pairs=set()
    for c in cards:
        if not isinstance(c,dict) or not isinstance(c.get('id'),str) or not isinstance(c.get('word'),str):raise ValueError('카드 ID/단어가 잘못되었습니다.')
        if not (1<=len(c['id'])<=90 and 1<=len(c['word'])<=90):raise ValueError('카드 ID/단어 길이 오류')
        if c['id'] in ids:raise ValueError('중복 카드 ID')
        ids.add(c['id']);pair=(c['word'].lower().strip(),str(c.get('pos','')).strip(),str(c.get('meaning','')).strip().casefold())
        if pair in pairs:raise ValueError('중복 단어/품사/뜻')
        pairs.add(pair)
        if c.get('status') not in ('draft','review','approved','published','rejected'):raise ValueError('잘못된 카드 상태')
        # Old approved/published cards are grandfathered only when unchanged.
        # Newly published or edited approved/published cards must pass all checks.
        if c['status'] in ('approved','published') and old_by_id.get(c['id']) != c:
            # Existing published items may be updated by the authenticated admin after explicit UI confirmation.
            # New publications and previously unpublished items retain the full approval gate.
            prior=old_by_id.get(c['id']) or {}
            if c['status']=='published' and prior.get('status')=='published':
                if CLOUD_ENABLED and not (str(c.get('sceneImage','')).startswith('assets/') or str(c.get('sceneImage','')).startswith('/api/public-image?')):
                    raise ValueError(c['word']+' 이미지 경로가 올바르지 않습니다.')
                if not all(c.get(k) for k in ('word','pos','ipa','meaning','mnemonic','scene','sceneImage')):
                    raise ValueError(c['word']+' 필수 공개 항목 누락')
                continue
            if CLOUD_ENABLED and not (str(c.get('sceneImage','')).startswith('assets/') or str(c.get('sceneImage','')).startswith('/api/public-image?')):
                raise ValueError(c['word']+' 이미지 경로가 올바르지 않습니다.')
            # Admin final approval: quality score and historical flags are advisory only.
            # Essential fields, stress range and clean image remain mandatory.
            stress=c.get('stress',[])
            if (not all(str(c.get(k,'')).strip() for k in ('word','pos','ipa','meaning','mnemonic','scene','example','translation','sceneImage'))
                or c.get('sceneImageMode')!='clean'
                or not isinstance(stress,list) or len(stress)!=2 or not all(isinstance(x,int) for x in stress)
                or stress[0]<0 or stress[1]>len(c['word']) or stress[1]<=stress[0]):
                raise ValueError(c['word']+' 공개 필수정보·강세·이미지 조건 미충족')
    return cards


def cloud_get_cards(admin):
    d=supa_request('/rest/v1/rpc/onecut_get_cards',method='POST',data={},service=True)
    if not isinstance(d,dict) or not isinstance(d.get('cards'),list):raise ValueError('DB 설치 또는 초기화 상태를 확인하세요.')
    return d


def cloud_save_cards(cards,version,admin):
    # Compare with the CURRENT cloud snapshot; never trust the client to claim
    # that an already published card was unchanged.
    snapshot=cloud_get_cards(admin)
    if snapshot.get('version') != version:
        raise ValueError('클라우드 버전이 변경됐습니다. 클라우드 새로고침 후 다시 시도하세요.')
    previous=snapshot['cards']
    validate_cards(cards, previous)
    cloud_backup_assets(cards, previous)
    return supa_request('/rest/v1/rpc/onecut_save_cards',method='POST',
      data={'p_cards':cards,'p_expected_version':version,'p_actor':admin['id']},service=True)


def validate_migration_cards(cards):
    """Validate shape and identity without applying *new-publication* quality gates.
    Historical approval / rejection / published fields are preserved byte-for-byte in JSON.
    """
    if not isinstance(cards,list) or not (1 <= len(cards) <= 3000):
        raise ValueError('이전할 카드 목록이 비어 있거나 올바르지 않습니다.')
    seen_ids=set(); seen_words=set()
    for card in cards:
        if not isinstance(card,dict) or not isinstance(card.get('id'),str) or not isinstance(card.get('word'),str):
            raise ValueError('기존 카드의 ID 또는 단어가 잘못되었습니다.')
        if not (1 <= len(card['id']) <= 90 and 1 <= len(card['word']) <= 90):
            raise ValueError('기존 카드 ID 또는 단어 길이가 잘못되었습니다.')
        pair=(card['word'].strip().casefold(),str(card.get('pos','')).strip(),str(card.get('meaning','')).strip().casefold())
        if card['id'] in seen_ids or pair in seen_words:
            raise ValueError('기존 카드에 중복된 ID 또는 단어/품사/뜻이 있습니다.')
        seen_ids.add(card['id']);seen_words.add(pair)
        if card.get('status') not in ('draft','review','approved','published','rejected'):
            raise ValueError(card['word']+'의 이전 상태가 올바르지 않습니다.')
    return cards


def cloud_migrate_legacy(cards,version,admin):
    # One-time migration only. The database RPC enforces *empty cloud state* atomically.
    validate_migration_cards(cards)
    missing=[];uploaded=0;seen=set()
    for card in cards:
        paths=[card.get('sceneImage','')]+[v.get('path','') for v in card.get('sceneCandidates',[]) if isinstance(v,dict)]
        for rel in paths:
            if not isinstance(rel,str) or not rel.startswith('assets/') or rel in seen:continue
            seen.add(rel)
            try:
                path=asset_path(rel)
                if not path.is_file():missing.append(rel);continue
                cloud_upload_image(rel);uploaded+=1
            except ValueError as exc:
                raise ValueError('이미지 업로드 오류 ('+rel+'): '+str(exc)) from exc
    data=supa_request('/rest/v1/rpc/onecut_migrate_legacy_cards',method='POST',
       data={'p_cards':cards,'p_expected_version':version,'p_actor':admin['id']},service=True)
    data['uploaded_images']=uploaded
    data['missing_images']=missing
    return data


def cloud_published():
    # Server queries database function with published-only, sanitized columns.
    published=supa_request('/rest/v1/rpc/onecut_public_cards',method='POST',data={},service=True)
    result=published if isinstance(published,list) else []
    for card in result:card['sceneImage']=cloud_public_image_path(card.get('sceneImage',''))
    return {'cards':result}


def asset_path(local_rel):
    if not isinstance(local_rel,str) or not local_rel.startswith('assets/') or '?' in local_rel or '#' in local_rel:
        raise ValueError('이미지 경로가 올바르지 않습니다.')
    target=(ROOT/local_rel).resolve()
    if not target.is_relative_to((ROOT/'assets').resolve()):raise ValueError('허용되지 않은 이미지 경로')
    return target


def cloud_upload_image(local_rel):
    """Back up generated/uploaded image to PRIVATE vocab-images bucket.
    Keep relative path in card JSON for existing local preview compatibility.
    Published cards resolve via /api/public-image (permission checked server side).
    """
    path=asset_path(local_rel)
    if path.suffix.lower() not in ('.png','.jpg','.jpeg','.webp') or not path.is_file():
        raise ValueError('실제 이미지 파일이 없습니다: '+local_rel)
    data=path.read_bytes()
    if len(data)>20_000_000:raise ValueError('이미지 파일이 너무 큽니다.')
    name='scenes/'+local_rel[len('assets/'):]
    mime=mimetypes.guess_type(str(path))[0] or 'image/png'
    # Upsert lets users resume cloud uploads after interrupted requests.
    supa_request('/storage/v1/object/vocab-images/'+quote(name,safe='/'),method='POST',data=data,raw=True,
        headers={'Content-Type':mime,'x-upsert':'true'},service=True,timeout=90)
    return local_rel


def cloud_backup_assets(cards, previous_cards=None):
    old_by_id={c.get('id'):c for c in (previous_cards or []) if isinstance(c,dict)}
    seen=set()
    for c in cards:
        values=[c.get('sceneImage','')]+[v.get('path','') for v in c.get('sceneCandidates',[]) if isinstance(v,dict)]
        for rel in values:
            if not isinstance(rel,str) or not rel.startswith('assets/') or rel in seen:continue
            seen.add(rel)
            path=asset_path(rel)
            # A candidate image that only existed on a previous computer can be
            # retained in DB, while published images MUST exist or be recoverable.
            if not path.is_file():
                if (c.get('status') in ('approved','published') and c.get('sceneImage')==rel
                        and old_by_id.get(c.get('id')) != c):
                    # A published image may have been uploaded to private Storage
                    # by an earlier version even if this version has no local copy.
                    try: cloud_image_bytes(rel)
                    except (ValueError, URLError):
                        raise ValueError(c.get('word','')+' 공개 이미지 파일이 PC 및 클라우드에 없습니다: '+rel)
                continue
            cloud_upload_image(rel)


def cloud_public_image_path(rel):
    if isinstance(rel,str) and rel.startswith('assets/'):
        return '/api/public-image?path='+quote(rel,safe='')
    return rel


def cloud_admin_image_urls(paths):
    """Short-lived, browser-usable links to PRIVATE images for an authenticated admin.
    Missing cloud assets are omitted; original card data is never modified.
    """
    if not isinstance(paths, list) or len(paths) > 60:
        raise ValueError('이미지 요청 개수가 올바르지 않습니다.')
    urls = {}
    for rel in dict.fromkeys(paths):
        if not isinstance(rel, str) or len(rel) > 500 or not rel.startswith('assets/'):
            continue
        asset_path(rel)
        obj = 'scenes/' + rel[len('assets/'):]
        try:
            signed = supa_request('/storage/v1/object/sign/vocab-images/' + quote(obj, safe='/'),
                                  method='POST', data={'expiresIn': 900}, service=True, timeout=18)
            url = signed.get('signedURL', '') if isinstance(signed, dict) else ''
            if url.startswith('/storage/v1/'):
                urls[rel] = SUPABASE_URL + url
            elif url.startswith('/object/'):
                urls[rel] = SUPABASE_URL + '/storage/v1' + url
            elif url.startswith(SUPABASE_URL + '/storage/v1/'):
                urls[rel] = url
        except ValueError:
            # Legacy paths can legitimately reference missing images.
            continue
    return {'urls': urls}


def cloud_image_bytes(rel):
    asset_path(rel) # validate path
    remote_path='scenes/'+rel[len('assets/'):]
    # Download from private bucket; never send service credential to browser.
    key=SUPABASE_SERVICE_KEY
    req=Request(SUPABASE_URL+'/storage/v1/object/vocab-images/'+quote(remote_path,safe='/'),
        headers={'apikey':key,'Authorization':'Bearer '+key},method='GET')
    try:
        with urlopen(req,timeout=30) as resp:return resp.read(),resp.headers.get('Content-Type') or 'image/png'
    except HTTPError as exc:raise ValueError('클라우드 이미지 조회 실패: HTTP '+str(exc.code)) from exc


def api_post(endpoint, payload, timeout=100):
    key = os.environ.get('OPENAI_API_KEY', '').strip()
    if not key:
        raise ValueError('OPENAI_API_KEY 환경변수가 설정되지 않았습니다. 안내서에 따라 설정해 주세요.')
    req = Request('https://api.openai.com/v1/' + endpoint,
                  data=json.dumps(payload, ensure_ascii=False).encode('utf-8'),
                  headers={'Authorization': 'Bearer ' + key,
                           'Content-Type': 'application/json'}, method='POST')
    try:
        with urlopen(req, timeout=timeout) as response:
            return json.load(response)
    except HTTPError as exc:
        try:
            msg = json.loads(exc.read().decode('utf-8')).get('error', {}).get('message', '')
        except Exception:
            msg = ''
        raise ValueError(f'AI API 응답 오류 ({exc.code}): {msg[:240] or exc.reason}') from exc
    except URLError as exc:
        raise ValueError(f'AI API에 연결하지 못했습니다: {exc.reason}') from exc


def require_card(data):
    if not isinstance(data, dict):
        raise ValueError('요청 형식이 잘못되었습니다.')
    for key in ('word', 'pos', 'ipa', 'meaning'):
        if not isinstance(data.get(key), str) or not data[key].strip():
            raise ValueError(f'필수 항목 누락: {key}')
    return data




RESEARCH_USER_PATH = ROOT.parent / 'onecut.research.json'
RESEARCH_BUILTIN_PATH = ROOT / 'research' / 'mnemonic_examples.json'

def _load_json_file(path, default):
    try:
        if path.is_file():
            obj=json.loads(path.read_text(encoding='utf-8'))
            return obj
    except Exception:
        pass
    return default

def _safe_research_item(item):
    if not isinstance(item,dict): return None
    verdict=str(item.get('verdict','')).strip().lower()
    if verdict not in ('success','failure'): return None
    word=str(item.get('word','')).strip()[:90]
    mnemonic=str(item.get('mnemonic','')).strip()[:260]
    if not word or not mnemonic: return None
    tags=item.get('tags',[])
    if not isinstance(tags,list): tags=[]
    return {
      'verdict':verdict,'word':word,'pos':str(item.get('pos',''))[:40],
      'ipa':str(item.get('ipa',''))[:120],'meaning':str(item.get('meaning',''))[:220],
      'anchor':str(item.get('anchor',''))[:140],'mnemonic':mnemonic,
      'scene':str(item.get('scene',''))[:500],'principle':str(item.get('principle') or item.get('reason',''))[:420],
      'tags':[str(x)[:50] for x in tags[:8] if str(x).strip()],
      'date':str(item.get('date',''))[:50]
    }

def load_research_examples(data=None, limit_each=5):
    """Return small, bounded examples for prompt grounding.

    Same-word examples come first, followed by recent/general examples. This is
    deliberately retrieval-like prompt context, not model fine-tuning.
    """
    built=_load_json_file(RESEARCH_BUILTIN_PATH,{'examples':[]})
    user=_load_json_file(RESEARCH_USER_PATH,{'examples':[]})
    pool=[]
    for src in (built.get('examples',[]) if isinstance(built,dict) else []):
        v=_safe_research_item(src)
        if v: pool.append(v)
    for src in (user.get('examples',[]) if isinstance(user,dict) else []):
        v=_safe_research_item(src)
        if v: pool.append(v)
    target=str((data or {}).get('word','')).strip().casefold()
    def rank(x):
        same=1 if target and x['word'].casefold()==target else 0
        return (same, x.get('date',''))
    successes=sorted([x for x in pool if x['verdict']=='success'],key=rank,reverse=True)[:limit_each]
    failures=sorted([x for x in pool if x['verdict']=='failure'],key=rank,reverse=True)[:limit_each]
    # Strip non-prompt metadata and keep context compact.
    fields=('word','ipa','meaning','anchor','mnemonic','principle','tags')
    return {
      'successes':[{k:x[k] for k in fields if x.get(k)} for x in successes],
      'failures':[{k:x[k] for k in fields if x.get(k)} for x in failures]
    }

def record_research_feedback(data):
    """Persist explicit admin judgment outside version folders.

    This lets a published result become a success example and a rejected or
    regenerated result become a failure example for later prompt retrieval.
    """
    if not isinstance(data,dict): raise ValueError('연구 피드백 형식 오류')
    item=_safe_research_item(data)
    if not item: raise ValueError('연구 피드백 필수값 누락')
    item['date']=datetime.now(timezone.utc).isoformat()
    store=_load_json_file(RESEARCH_USER_PATH,{'version':1,'examples':[]})
    if not isinstance(store,dict): store={'version':1,'examples':[]}
    arr=store.get('examples',[])
    if not isinstance(arr,list): arr=[]
    # Avoid identical consecutive duplicates while preserving admin history.
    sig=(item['verdict'],item['word'].casefold(),item['mnemonic'].strip().casefold())
    cleaned=[]
    for old in arr[-399:]:
        v=_safe_research_item(old)
        if v: cleaned.append(old)
    if cleaned:
        last=_safe_research_item(cleaned[-1])
        last_sig=(last['verdict'],last['word'].casefold(),last['mnemonic'].strip().casefold()) if last else None
    else:last_sig=None
    if sig!=last_sig: cleaned.append(item)
    store={'version':1,'updated_at':item['date'],'examples':cleaned[-400:]}
    tmp=RESEARCH_USER_PATH.with_suffix('.json.tmp')
    tmp.write_text(json.dumps(store,ensure_ascii=False,indent=2),encoding='utf-8')
    tmp.replace(RESEARCH_USER_PATH)
    return {'saved':True,'counts':{
      'success':sum(1 for x in store['examples'] if isinstance(x,dict) and x.get('verdict')=='success'),
      'failure':sum(1 for x in store['examples'] if isinstance(x,dict) and x.get('verdict')=='failure')}}

def research_phonetic_anchors(data):
    """Research natural Korean sound anchors before writing mnemonic sentences.

    This stage deliberately does not write the final mnemonic. It searches for
    real Korean words / common loanwords / short everyday expressions that are
    genuinely close to the IPA and can bridge directly to the English meaning.
    """
    data = require_card(data)
    system = (
        "당신은 한국어 발음연상 영단어의 '발음 고리 연구원'입니다. 최종 연상문장을 쓰지 말고 먼저 소리 고리만 연구하세요. "
        "철자가 아니라 IPA와 강세를 기준으로 영어 소리를 듣는다고 가정하세요. 한국인이 실제로 쓸 수 있는 한국어 단어, 널리 쓰는 외래어, 짧은 대사 및 스토리 속에서 자연스럽게 소개되는 인명·지명도 후보로 허용합니다. 인명은 대사 속에서 누구인지 알 수 있어야 하며 무작위 음절 조합은 허용하지 않습니다. "
        "의미 없는 감탄사(어~, 억!, 툭!), 존재하지 않는 합성어, 억지 조사 붙이기, 철자만 보고 만든 말은 금지합니다. "
        "각 후보는 (a) IPA의 어느 구간과 닮는지, (b) 실제 한국어에서 어떻게 쓰이는지, (c) 그 표현의 실제 기능/상황이 영어 뜻과 직접 이어지는지를 따로 검증하세요. "
        "발음만 닮고 뜻 연결이 없거나, 뜻만 맞고 발음 고리가 없으면 reject입니다. 다만 소리 고리 자체에 뜻이 없어도 후속 장면에서 원인-결과로 뜻을 명확히 연결할 수 있다면 storyBridge로 인정하세요. 일부 음절만 닮을 때는 coverage를 낮게 주세요. "
        "성공 원리 예시: conceal /kənˈsiːl/ ↔ '컨실러'처럼 실제 쓰이는 외래어가 있고, 컨실러의 실제 기능인 '잡티를 가리다'가 conceal의 '숨기다/감추다'와 직접 이어지는 경우. "
        "실패 예시: conceal ↔ '가방 안에 몰래 숨겼어'는 뜻만 맞고 영어 소리 고리가 없으므로 reject. abandon ↔ '네가 밴 돈이잖아'는 일부 소리만 닮고 버리다/포기하다와 연결되지 않아 reject. "
        "외래어가 원 영어와 어원이 닿아 있어도 학습자에게 익숙하고 발음·기능·뜻이 동시에 이어지면 기억 고리로 사용할 수 있습니다. 다만 모든 단어에 외래어를 억지 적용하지 마세요. "
        "좋은 고리가 없으면 없다고 판정하는 것이 정답입니다. 억지로 6개를 채우지 마세요. "
        "researchExamples가 있으면 success는 원리를 참고하고 failure는 같은 실패 구조를 피하세요. 문장 자체를 복제하지 말고 소리-뜻-장면 연결 원리만 일반화하세요. "
        "오직 JSON: {\"anchors\":[{\"anchor\":\"실제 표현\",\"ipaSpan\":\"/..../\",\"coverage\":0,\"naturalUse\":\"실제 쓰임\",\"meaningBridge\":\"뜻 연결\",\"verdict\":\"pass 또는 reject\",\"reason\":\"판정 근거\"}],\"researchSummary\":\"...\"}. "
        "coverage는 0~100이며 pass는 원발음의 핵심 강세부를 충분히 포착하고 실제 쓰임과 뜻 연결까지 모두 만족할 때만 허용합니다."
    )
    payload={k:data.get(k,'') for k in ('word','pos','ipa','hangul','meaning')}
    payload['researchExamples']=load_research_examples(data)
    result=api_post('chat/completions',{'model':AI_MODEL,'temperature':0.35,
      'response_format':{'type':'json_object'},'messages':[{'role':'system','content':system},{'role':'user','content':json.dumps(payload,ensure_ascii=False)}]})
    obj=json.loads(result['choices'][0]['message']['content'])
    anchors=obj.get('anchors',[])
    if not isinstance(anchors,list):
        anchors=[]
    cleaned=[]
    for a in anchors[:8]:
        if not isinstance(a,dict):
            continue
        anchor=str(a.get('anchor','')).strip()[:120]
        if not anchor:
            continue
        try: coverage=int(round(float(a.get('coverage',0))))
        except Exception: coverage=0
        coverage=max(0,min(100,coverage))
        verdict=str(a.get('verdict','reject')).strip().lower()
        if verdict not in ('pass','reject'): verdict='reject'
        # Fail closed: weak phonetic coverage cannot be marked pass.
        if coverage < 55: verdict='reject'
        cleaned.append({
          'anchor':anchor,'ipaSpan':str(a.get('ipaSpan',''))[:80],
          'coverage':coverage,'naturalUse':str(a.get('naturalUse',''))[:320],
          'meaningBridge':str(a.get('meaningBridge',''))[:320],
          'verdict':verdict,'reason':str(a.get('reason',''))[:420]
        })
    return {'anchors':cleaned,'researchSummary':str(obj.get('researchSummary',''))[:700],
            'model':AI_MODEL,'method':'phonetic-anchor-research-v514'}

def generate_mnemonics(data):
    data = require_card(data)
    anchor_research = data.get('anchorResearch') if isinstance(data.get('anchorResearch'), dict) else None
    if not anchor_research:
        anchor_research = research_phonetic_anchors(data)
    passed_anchors = [a for a in anchor_research.get('anchors', []) if isinstance(a, dict) and a.get('verdict') == 'pass']
    system = (
        "당신은 한국어 발음연상 영어단어 학습 콘텐츠를 개발하는 작가입니다. "
        "절대 규칙: 한국어에서 실제 쓰이는 단어·구어체만 사용하세요. '어~', '억!', '툭!' 등 맥락 없는 감탄사·의성어로 발음을 억지로 메우지 마세요. "
        "단어 철자보다 정확한 IPA의 음가와 강세를 먼저 읽으세요. 영어 발음에서 자연스럽게 닮는 한국어 구절을 찾되, 소리 일부만 일치하면 전체가 비슷하다고 과장하지 마세요. "
        "사고 순서는 (1) 영어 소리와 닮은 대사·인명·외래어를 폭넓게 찾기, (2) 뜻을 행동이나 반응으로 보여주는 상황 찾기, (3) 소리와 뜻을 인과관계로 연결하기, (4) 한 문장과 한 장면으로 압축하기입니다. 단순 뜻 묘사와 억지 발음 끼워 맞추기는 불합격입니다. "
        "영어 뜻이 장면 속 행동이나 결과 자체로 드러나야 하며, 연상 문장만 봐도 뜻을 복원할 수 있어야 합니다. 대사만 비슷하고 뜻과 다른 상황이면 실패입니다. "
        "명사·동사 등 품사의 의미를 혼동하지 마세요. 한국어 문법이 부자연스럽거나 한 컷으로 바로 표현할 수 없으면 실패입니다. "
        "사용 가능한 연상 고리는 자연스러운 한국어 대사, 친숙한 일반명사, 널리 쓰이는 외래어, 이야기 안에서 인물로 명확히 설정되는 이름도 포함합니다. 이름은 발음만 맞추기 위한 무의미한 음절 조합이면 안 되며 상황과 대사 속에서 자연스러워야 합니다. "
        "발음 유사성과 뜻 직접 연결을 동시에 만족하는 후보를 찾지 못했다면 꾸며내서 good이라고 하지 말고 weak로 표시하세요. '무리한 연상보다 보류'가 원칙입니다. "
        '반드시 최초에 영어 IPA의 주요 소리·강세에 대응하는 실제 한국어/일상 외래어/상용어 후보를 먼저 탐색하세요. 음성 고리(anchor)는 사전에 있는 자연스러운 단어, 일반적으로 쓰는 외래어, 자연스러운 짧은 대사여야 합니다. 가능한 경우 익숙한 외래어의 발음과 기능을 함께 이용해 보세요. 대표적인 분석 예시: conceal /kənˈsiːl/ (감추다) ↔ 컨실러(잡티를 가리는 화장품) ↔ 거울 앞에서 “컨실러로 잡티를 싹 숨겼어!”라는 자연스러운 대사. 이 예시는 원리가 성공적인 대비 사례일 뿐 정답을 복제하거나 모든 단어에 외래어를 억지 적용하지 마세요. 반례: conceal에 “가방 안에 몰래 숨겼어”, “책 뒤에 숨겼어”, “비밀을 잘 숨겼네”는 뜻 설명은 되지만 영어 소리와의 연결고리가 없어 불합격입니다. 더 나쁜 반례: abandon에 “네가 밴 돈이잖아! 돌려줘!”는 발음 일부는 닮아도 버리다/포기하다로 이어지지 않으므로 불합격입니다. 후보마다 anchor에 실제 사용하는 단어(예: 컨실러), IPA 대응 음절 및 남는 차이를 명시하고, reason에 그 단어가 쓰이는 실제 기능이 영어 뜻과 직결되는 행동을 명시하세요. 실제 한국어에서 사용하는 외래어가 영어 타깃 단어에서 파생된 경우에도 기억 단서로 허용하되, 발음 일치·뜻 직결이 없으면 불합격입니다. '
        "후보마다 직접 발음형·스토리 연상형·의미 결합형을 고르게 시도하고, 각각 다른 장면과 소리 고리를 사용하세요. 예시: admire → 어디, 마이어? → 마이어의 작품을 찾아 감탄하는 장면. 단, 이 예시를 다른 단어에 복제하지 말고 발음 차이도 정직하게 평가하세요. "
        "후보마다 각각 다른 장면과 소리 고리를 사용하세요. 생성된 자기평가는 참고일 뿐이며, 실제 발음과 일상 언어가 다르면 점수는 낮아야 합니다. "
        "priorFailures에 제시된 실패 문장·이유를 학습해 유사한 억지 구조를 되풀이하지 마세요. "
        "researchExamples의 success/failure 사례도 참고하되 표현을 베끼지 말고, 승인된 사례의 연결 원리와 반려 사례의 실패 원인을 새로운 단어에 일반화하세요. "
        "anchorResearch의 passedAnchors가 있으면 그 고리를 우선 재료로 사용하되 그대로 복사하지 말고 자연스러운 실제 대사와 장면으로 발전시키세요. passedAnchors가 비어 있으면 발음 고리가 약하다는 뜻이므로 뜻만 묘사한 문장을 good으로 만들지 마세요. 필요하면 세 후보 모두 weak로 보류하세요. "
        "경선식 등 타 교재의 문장이나 장면은 복제하지 말고 완전히 독창적인 콘텐츠를 작성하세요. "
        "한 컷 장면의 등장인물·행동·장소를 구체적으로 서술하고, 시각적으로 전달할 핵심 대사를 명시하세요. "
        "오직 JSON 객체로 답하십시오. 형식: "
        '{"candidates":[{"mnemonic":"자연스럽고 짧은 한국어 대사 한 문장","scene":"구체적이고 현실적인 한 컷 장면",'
        '"anchor":"IPA 실제 소리↔한국어 소리, 일치하는 구간과 일치하지 않는 구간을 솔직히 구분",'
        '"reason":"발음이 닮는 근거, 상황의 행동이 뜻에 해당하는 근거",'
        '"quality":"good 또는 weak","selfScores":{"sound":0,"natural":0,"meaning":0,"visual":0}}, ...]}. '
        "selfScores 배점은 sound 30, natural 25, meaning 25, visual 20. 높은 점수라도 소리 왜곡·뜻 반전·억지 감탄사·존재하지 않는 표현이 있으면 quality=weak입니다. "
        "반드시 후보 3개를 반환하되 부적합하면 모두 weak도 허용합니다. 영어 철자·IPA·품사·뜻은 변경하지 마세요."
    )
    user_data = {k: data.get(k, '') for k in ('word','pos','ipa','hangul','meaning','mnemonic','scene','reason')}
    user_data['anchorResearch'] = {
        'passedAnchors': passed_anchors[:5],
        'summary': str(anchor_research.get('researchSummary',''))[:500]
    }
    user_data['researchExamples']=load_research_examples(data)
    # Feed only brief, bounded failure evidence from previous rounds; never echo private data.
    failures = data.get('priorFailures', [])
    if isinstance(failures, list):
        user_data['priorFailures'] = [
            {'mnemonic':str(f.get('mnemonic',''))[:180],
             'reason':str(f.get('reason',''))[:300]}
            for f in failures[:9] if isinstance(f, dict)
        ]
    user = json.dumps(user_data, ensure_ascii=False)
    result = api_post('chat/completions', {
        'model': AI_MODEL, 'temperature': 0.72,
        'response_format': {'type': 'json_object'},
        'messages': [{'role': 'system', 'content': system}, {'role': 'user', 'content': user}]
    })
    raw = result['choices'][0]['message']['content']
    parsed = json.loads(raw)
    candidates = parsed.get('candidates', [])
    if not isinstance(candidates, list) or len(candidates) != 3:
        raise ValueError('AI가 후보 3개를 반환하지 않았습니다. 다시 생성해 주세요.')
    cleaned = []
    limits=(30,25,25,20)
    for item in candidates:
        if not isinstance(item, dict) or not str(item.get('mnemonic', '')).strip():
            raise ValueError('AI 응답 형식이 올바르지 않습니다.')
        scores=item.get('selfScores') if isinstance(item.get('selfScores'),dict) else {}
        safe_scores={}
        for key,limit in zip(('sound','natural','meaning','visual'),limits):
            try:value=int(round(float(scores.get(key,0))))
            except Exception:value=0
            safe_scores[key]=max(0,min(limit,value))
        quality=str(item.get('quality','weak')).strip().lower()
        if quality not in ('good','weak'):quality='weak'
        mnemonic_text=str(item.get('mnemonic','')).strip()
        anchor_text=str(item.get('anchor','')).strip()
        # 모델 자기평가가 낮은데 good으로 낙관하는 것을 방지한다.
        if sum(safe_scores.values()) < 80 or safe_scores['sound'] < 21 or safe_scores['meaning'] < 18:
            quality='weak'
        # Research-first fail closed: a good candidate must be grounded in at least one
        # independently researched pass anchor, and the claimed anchor must actually
        # appear in the learner-facing Korean mnemonic (spaces/punctuation ignored).
        def _norm_anchor(v):
            return re.sub(r'[^0-9A-Za-z가-힣]+','',str(v)).lower()
        researched=[_norm_anchor(a.get('anchor','')) for a in passed_anchors if isinstance(a,dict)]
        norm_anchor=_norm_anchor(anchor_text); norm_mnemonic=_norm_anchor(mnemonic_text)
        grounded=bool(any(x and x in norm_mnemonic for x in researched))
        spoken=grounded
        if not passed_anchors or not grounded or not spoken:
            quality='weak'
        cleaned.append({
            'mnemonic':mnemonic_text[:500],
            'scene':str(item.get('scene',''))[:900],
            'anchor':anchor_text[:240],
            'reason':str(item.get('reason',''))[:900],
            'quality':quality,
            'selfScores':safe_scores
        })
    return {'candidates': cleaned, 'anchorResearch': anchor_research, 'model': AI_MODEL, 'method': 'story-aware-mnemonic-v514'}



def evaluate_mnemonic(data):
    """AI recommendation, not a validated pronunciation dictionary. Human verification remains mandatory."""
    data=require_card(data)
    if not str(data.get('mnemonic','')).strip() or not str(data.get('scene','')).strip():
        raise ValueError('연상문장과 장면을 먼저 작성해 주세요.')
    system=(
        '당신은 한국어 발음연상 영단어 교재의 엄격한 편집 심사위원입니다. '
        '실제 영어 발음(IPA)과 한국어 대사의 닮음이 전체 발음·강세 기준으로 얼마나 충분한지, 한국어 실제 일상표현인지, 문장만 떠올려도 뜻이 바로 회상되는지, 한 컷의 행동 자체가 뜻인지 별도로 평가하세요. 뜻이 장면에 없으면 반드시 meaning 점수를 10 이하로 제한하세요. '
        'IPA 자체를 외부 사전에서 조회한 것처럼 주장하지 마세요. IPA에 의문이 있다면 soundCheckNeeded=true. '
        '뜻만 맞는 평범한 문장은 발음연상으로 인정하지 마세요. 반례: conceal에서 “가방 안에 몰래 숨겼어”, “책 뒤에 숨겼어”, “비밀을 잘 숨겼네”는 영어 소리 연결이 없으므로 sound 점수를 5점 이하로 제한하세요. 긍정 참고: conceal ↔ 컨실러(상용 화장품) ↔ 잡티를 가리는 실제 행동은 발음 및 뜻 연결의 유력 후보입니다. 단, 긍정 사례도 IPA 음가 차이를 검증하고 실제 문장이 자연스러운지 다시 확인하세요. anchor가 제시되어도 말 속에 그 소리 고리가 실제로 등장하지 않으면 sound 점수를 10점 이하로 제한하세요. 요청의 anchor 필드가 있으면 그 표현이 mnemonic 대사에 실제로 포함되는지 확인하고, 포함되지 않으면 sound를 10 이하로 제한하세요. '
        '의미 없는 어~ 억~ 툭 등의 감탄사, 일상에서 쓰지 않는 합성어, 뒤늦은 사후 해석으로 뜻을 꾸며낸 경우 unnatural=true입니다. IPA에서 일부 소리만 겨우 일치하면 sound 점수 15 이하로 제한하세요. 실제 대사로서 말이 안 되면 natural 점수 10 이하로 제한하세요. 그림 실물을 보지 못하므로 장면 설명의 표현 가능성만 평가합니다. 자체 평가가 90점대여도 근거가 빈약한 후보는 낮게 점수를 주세요. 후보의 quality 태그를 신뢰하지 말고 독립 평가하세요. researchExamples가 있으면 성공 사례의 원리와 실패 사례의 경고를 독립 검수 참고자료로만 사용하고 문구 자체는 정답으로 취급하지 마세요. '
        '선행 콘텐츠나 교재의 고유 문장은 복제하지 마세요. 소리와 뜻 연결이 실제로 없는 경우 총점 85 이상을 절대 부여하지 마세요. 소리 고리가 원발음의 짧은 조각에만 대응한다면 그 불일치를 명시하고 엄격하게 감점하세요. 의미 없는 음절 붙이기, 철자 기반 말장난, 어색한 조사/외래어 변형, 뜻과 반대되는 행동, 발음만 비슷하고 뜻 연결이 없는 연상은 크게 감점하세요. 한국인이 실제로 말할 법한 짧은 구어체와 장면의 원인-결과 연결을 높게 평가하세요. '
        '반드시 JSON 객체 하나만 답하세요. 형식: '
        '{"scores":[0,0,0,0],"reasons":["...","...","...","..."],'
        '"summary":"...","soundCheckNeeded":false,"unnatural":false,"speechBubbleRequired":true,'
        '"suggestedSpeech":"짧은 한국어 핵심 대사 또는 빈 문자열"}. '
        '점수 상한은 순서대로 30,25,25,20. 이유는 한국어 한두 문장.'
    )
    payload={k:data.get(k,'') for k in ('word','pos','ipa','hangul','meaning','mnemonic','scene','anchor')}
    payload['researchExamples']=load_research_examples(data)
    result=api_post('chat/completions',{'model':AI_MODEL,'temperature':0.2,
      'response_format':{'type':'json_object'},'messages':[{'role':'system','content':system},
      {'role':'user','content':json.dumps(payload,ensure_ascii=False)}]})
    obj=json.loads(result['choices'][0]['message']['content'])
    scores=obj.get('scores');reasons=obj.get('reasons')
    limits=(30,25,25,20)
    if not isinstance(scores,list) or len(scores)!=4 or any(type(x) not in (int,float) or x<0 or x>limits[i] for i,x in enumerate(scores)):
        raise ValueError('AI 평가 점수 형식이 잘못되었습니다. 다시 시도해 주세요.')
    if not isinstance(reasons,list) or len(reasons)!=4:raise ValueError('AI 평가 근거가 누락되었습니다.')
    return {'scores':[int(round(x)) for x in scores], 'reasons':[str(x)[:320] for x in reasons],
      'summary':str(obj.get('summary',''))[:700], 'soundCheckNeeded':obj.get('soundCheckNeeded') is True,
      'unnatural':obj.get('unnatural') is True, 'speechBubbleRequired':obj.get('speechBubbleRequired') is True,
      'suggestedSpeech':str(obj.get('suggestedSpeech',''))[:160], 'model':AI_MODEL}

def verify_mnemonic_strict(data):
    """A separate skeptical review, used only for a high-scoring finalist.

    AI approval is advisory; this is intentionally fail-closed when uncertain.
    """
    data = require_card(data)
    if not str(data.get('mnemonic', '')).strip() or not str(data.get('scene', '')).strip():
        raise ValueError('연상문장과 장면이 모두 필요합니다.')
    system = (
        '당신은 AI 발음연상 콘텐츠를 심사하는 반대측 검수자입니다. 이전 점수나 추천을 신뢰하지 말고 탈락 사유를 찾으세요. '
        '원 영어 IPA와 한국어 연상문장 사이의 실제 음성 닮음, 실제 한국어 문장 자연스러움, 영어 뜻과 장면의 직접적인 연결을 독립적으로 따지세요. '
        '실제 한국어 연상 대사에 타깃 영어 단어와 닮은 발음 고리가 포함되는지 먼저 검증하세요. anchor 필드가 주어지면 그 실제 표현이 mnemonic 대사에 등장해야 하며, 등장하지 않으면 soundDirect=false입니다. 발음 연결 없이 뜻만 묘사하면 soundDirect=false입니다. conceal “가방 안에 몰래 숨겼어”는 불합격 예시이며, 컨실러가 잡티를 가린다는 장면은 긍정 후보입니다(모든 긍정 후보도 독립 검증 필수). 외래어 연상을 일률적으로 부적합하게 취급하지 마세요. researchExamples가 있으면 관리자 승인/반려 사례의 원리를 참고하되, 동일 문구를 자동 통과시키지 말고 현재 단어의 IPA·뜻에 대해 다시 검증하세요. '
        '이름을 활용한 이야기라도 그림 속 인물·행동·반응으로 뜻을 회상할 수 있다면 meaningDirect=true가 가능합니다. 뜻과 장면이 서로 관련 없어도 이야기로 우길 수 있는 경우는 meaningDirect=false. '
        '소리 조각 하나만 억지로 비슷하게 만든 경우는 soundDirect=false. '
        '그림에서 묘사되는 행동 자체가 뜻의 의미와 다르거나 반대라면 meaningDirect=false. '
        '한국인이 일상에서 절대 쓰지 않을 문장이라면 naturalKorean=false. '
        '실제 한국어에서 거의 사용하지 않는 말이나 맥락 없는 감탄사로 영어 발음의 빈 부분을 때운다면 naturalKorean=false, 소리 일부에 억지로 기대면 soundDirect=false. 장면을 자세히 해설해야만 뜻이 이해된다면 meaningDirect=false. 소리 연결, 뜻 연결이 확실하지 않다면 false로 보수적으로 판정하세요. '
        'JSON만 반환: {"soundDirect":false,"meaningDirect":false,"naturalKorean":false,'
        '"phoneticEvidence":"구체적인 실제 음절 대응", "meaningEvidence":"장면의 행동과 뜻이 연결되는 근거",'
        '"criticalIssue":"없으면 빈 문자열, 있으면 탈락 사유"}. '
        '외부 교재의 문장이나 그림을 복제하지 마세요.'
    )
    request = {k:data.get(k,'') for k in ('word','pos','ipa','hangul','meaning','mnemonic','scene','anchor')}
    request['researchExamples']=load_research_examples(data)
    result = api_post('chat/completions', {
        'model': AI_MODEL, 'temperature': 0.0, 'response_format': {'type':'json_object'},
        'messages':[{'role':'system','content':system},
                    {'role':'user','content':json.dumps(request,ensure_ascii=False)}]
    })
    parsed = json.loads(result['choices'][0]['message']['content'])
    # Only explicit booleans count as a pass. Missing data and uncertainty fail closed.
    pillars=('soundDirect','meaningDirect','naturalKorean')
    passed=all(parsed.get(k) is True for k in pillars) and not str(parsed.get('criticalIssue','')).strip()
    evidence={k:str(parsed.get(k,'')).strip()[:400] for k in ('phoneticEvidence','meaningEvidence','criticalIssue')}
    if not evidence['phoneticEvidence'] or not evidence['meaningEvidence']:passed=False
    return {'passed':passed,'checks':{k:parsed.get(k) is True for k in pillars},'evidence':evidence,'model':AI_MODEL}


def generate_image(data):
    data = require_card(data)
    if not data.get('mnemonic', '').strip() or not data.get('scene', '').strip():
        raise ValueError('승인된 연상문장과 장면 설명이 모두 필요합니다.')
    image_dir = ROOT / 'assets' / 'generated'
    image_dir.mkdir(parents=True, exist_ok=True)
    # 학습카드 제목/IPA/뜻/관리자 상태는 HTML에서만 표시한다.
    # 이미지에는 학습에 필요한 핵심 한국어 대사(및 scene에 명시된 필수 간판) 외의 읽을 수 있는 텍스트를 넣지 않는다.
    raw_feedback = str(data.get('imageFeedback', '') or data.get('reason', '') or '').strip()
    blocked_meta = ('AI 자동 제작','관리자','공개 대기','검수','품질 미달','재작업 대상','초안','status','score')
    safe_feedback = '' if any(x.lower() in raw_feedback.lower() for x in blocked_meta) else raw_feedback[:300]
    prompt = ("Create ONE original square (1:1) Korean educational cartoon SCENE, not a vocabulary poster. "
              "Premium modern Korean MANHWA/WEBTOON illustration aesthetic: slender naturally proportioned teenage or adult characters, fine expressive pencil-and-ink linework, subtle cool-grey and natural restrained colors, mature facial features and credible anatomy. NO childish chibi, mascot, cute infographic characters, thick cartoon linework, plastic 3D, clip art, bright yellow decorative ribbons or poster layouts. One clear central action, "
              "balanced square composition filling the frame without cropping characters, crucial objects, cafe sign or speech bubbles; keep clear 5 percent safe margins. Choose a distinctly different new protagonist for each DIFFERENT English word: vary age group, gender presentation, face shape, hairstyle, hair color within natural Korean looks, body type, wardrobe, occupation and viewpoint; do not reuse the same generic black-haired hoodie teenager across vocabulary entries. Make personality and appearance appropriate to the scenario, using children, teens, adults or seniors as appropriate. If regenerating the SAME word, preserve the basic protagonist and scene where possible, unless admin requests a change. Rotate gesture, framing and camera angle to avoid repetitive episodes. Keep background sparse and contextual; NO decorative leaves, flowers, plants, arbitrary sparkles, light rays or patterns. "
              "Draw only the objects needed for the user's concrete phonetic mnemonic, with expressive readable body language. Show the exact concrete meaning by the character action or result, not only by a caption. The scene should remain meaningful even if the dialogue balloon is hidden. Use realistic everyday interpersonal reactions instead of contrived symbolic compositions. "
              "No English word title, no IPA, no Korean definition box, no stress coloring, no footer mnemonic summary, "
              "no heading, no graphic frame, no promotional decorative shapes. Do not include emoji, graphic explanation overlays, or a heading. "
              "CRITICAL TEXT HYGIENE: never draw any admin/system/UI/status text, app interface, dashboard, progress label, score, button, watermark, model name, generation note, or production metadata. Never draw phrases such as AI 자동 제작, 관리자, 공개 대기, 검수, 초안, 점수, status. "
              "Do not invent readable text on phones, tablets, monitors, papers, packaging, mirrors, posters or background objects. The ONLY readable text allowed is (a) the exact sceneSpeech in one character speech/thought bubble and (b) an essential short in-scene sign ONLY when that sign is explicitly required by the scene description. "
              "If the mnemonic contains dialogue spoken or thought by a character, MUST draw the exact key Korean phrase IN A SPEECH BUBBLE — specifically ONE LARGE SPEECH BUBBLE — within the picture, not merely in a separate footer. "
              "The main speech bubble must be intentionally LARGE for smartphone learning: roughly 35–45% of the image width and at least about 22% of the image height when needed, with very large bold clean Korean lettering, high contrast, generous inner padding, and preferably no more than 2–3 short lines. The dialogue must remain comfortably readable when the full square image is displayed at about 320 CSS pixels wide on a phone. Keep the bubble fully inside a 7% safe margin and never clip it. "
              "If sceneSpeech is given, print that exact Korean speech and no extra sentence in the main bubble. Make every Korean syllable accurate. "
              "Do not imitate any existing course artwork or textbook layout. Clear and memorable on a smartphone.\n"
              f"English vocabulary word (context only, DO NOT print): {data['word']}\n"
              f"Meaning (context only, DO NOT print): {data['meaning']}\n"
              f"Natural Korean mnemonic: {data['mnemonic']}\n"
              f"Scene to draw faithfully: {data['scene']}\n"
              f"Exact speech bubble dialogue (if applicable): {data.get('sceneSpeech','')}\n"
              f"Optional visual correction feedback (context only; NEVER print): {safe_feedback}")
    result = api_post('images/generations', {'model': IMAGE_MODEL, 'prompt': prompt, 'size': '1024x1024', 'n': 1}, timeout=210)
    item = result['data'][0]
    if 'b64_json' not in item:
        raise ValueError('이미지 데이터가 반환되지 않았습니다. 이미지 모델 설정을 확인해 주세요.')
    data_bytes = base64.b64decode(item['b64_json'], validate=True)
    if not data_bytes.startswith(b'\x89PNG\r\n\x1a\n') or len(data_bytes) > 20_000_000:
        raise ValueError('PNG 이미지가 올바르지 않습니다.')
    slug = re.sub(r'[^a-zA-Z0-9_-]+','_',data['word']).strip('_')[:45] or 'card'
    stamp = datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S_%f')
    filename = f'{slug}_{stamp}.png'
    (image_dir / filename).write_bytes(data_bytes)
    local_path=f'assets/generated/{filename}'
    if CLOUD_ENABLED:
        try:cloud_upload_image(local_path)
        except Exception as exc:raise ValueError('클라우드 이미지 업로드 실패: '+str(exc)) from exc
    return {'image': local_path, 'model': IMAGE_MODEL}



def review_generated_image(data):
    """Vision-based QA for a generated mnemonic scene.

    This is a post-generation quality gate, not OCR-only validation. It checks
    for forbidden admin/UI text, speech-bubble readability, semantic scene fit,
    and composition. It never publishes or mutates a card.
    """
    data=require_card(data)
    rel=str(data.get('image','')).strip().replace('\\','/')
    if not rel.startswith('assets/generated/'):
        raise ValueError('검수 가능한 생성 이미지 경로가 아닙니다.')
    path=(ROOT/rel).resolve()
    allowed=(ROOT/'assets'/'generated').resolve()
    if allowed not in path.parents or not path.is_file():
        raise ValueError('검수할 이미지 파일을 찾을 수 없습니다.')
    raw=path.read_bytes()
    if len(raw)>20_000_000 or not raw.startswith(b'\x89PNG\r\n\x1a\n'):
        raise ValueError('검수할 PNG 이미지가 올바르지 않습니다.')
    encoded=base64.b64encode(raw).decode('ascii')
    expected=str(data.get('sceneSpeech','') or data.get('mnemonic','')).strip()[:220]
    system=(
        "당신은 모바일 영단어 웹툰 이미지의 엄격한 최종 시각 검수자입니다. 이미지 자체만 보고 판정하세요. "
        "다음은 즉시 불합격: AI/관리자/공개/검수/초안/점수/status/버튼/대시보드 등 시스템·UI 문구가 보임, "
        "학습과 무관한 읽을 수 있는 텍스트가 보임, 핵심 말풍선이 너무 작거나 잘려 320px 모바일에서 읽기 어려움, "
        "요청한 핵심 한국어 대사가 크게 다르거나 오탈자가 심함, 장면의 행동이 영어 뜻을 직접 보여주지 못함. "
        "말풍선은 대략 이미지 폭 30% 이상을 차지하는 충분히 큰 크기여야 하며 글씨는 굵고 명확해야 합니다. "
        "텍스트가 전혀 필요 없는 장면이면 speechReadability를 감점하지 마세요. "
        "오직 JSON: {\"passed\":true,\"scores\":{\"textHygiene\":0,\"speechReadability\":0,\"sceneMeaning\":0,\"composition\":0},"
        "\"issues\":[\"...\"],\"retryFeedback\":\"다음 이미지 생성에서 고칠 구체적 지시\"}. "
        "각 점수는 25점 만점. passed=true는 총점 85 이상이며 textHygiene>=23, speechReadability>=20, sceneMeaning>=20일 때만 허용합니다."
    )
    context=(f"단어: {data['word']}\n뜻: {data['meaning']}\n연상문장: {data.get('mnemonic','')}\n"
             f"장면 설명: {data.get('scene','')}\n기대 핵심 대사: {expected or '(대사 없음)'}")
    result=api_post('chat/completions',{
      'model':AI_MODEL,'temperature':0.1,'response_format':{'type':'json_object'},
      'messages':[{'role':'system','content':system},{'role':'user','content':[
        {'type':'text','text':context},
        {'type':'image_url','image_url':{'url':'data:image/png;base64,'+encoded,'detail':'low'}}
      ]}]},timeout=120)
    parsed=json.loads(result['choices'][0]['message']['content'])
    scores=parsed.get('scores') if isinstance(parsed.get('scores'),dict) else {}
    safe={}
    for k in ('textHygiene','speechReadability','sceneMeaning','composition'):
        try:v=int(round(float(scores.get(k,0))))
        except Exception:v=0
        safe[k]=max(0,min(25,v))
    total=sum(safe.values())
    issues=parsed.get('issues') if isinstance(parsed.get('issues'),list) else []
    issues=[str(x)[:220] for x in issues[:8] if str(x).strip()]
    passed=(parsed.get('passed') is True and total>=85 and safe['textHygiene']>=23 and safe['speechReadability']>=20 and safe['sceneMeaning']>=20)
    return {'passed':passed,'scores':safe,'total':total,'issues':issues,
            'retryFeedback':str(parsed.get('retryFeedback',''))[:500],'model':AI_MODEL,
            'method':'vision-image-qa-v454'}


def _decode_upload_base64(value, label='파일'):
    raw=str(value or '').strip()
    if not raw: raise ValueError(label+' 데이터가 없습니다.')
    if ',' in raw and raw.lower().startswith('data:'): raw=raw.split(',',1)[1]
    try: return base64.b64decode(raw,validate=True)
    except Exception as exc: raise ValueError(label+' 데이터를 읽을 수 없습니다.') from exc

def _store_scene_bytes(data, prefix='package'):
    if not (100 < len(data) <= 8*1024*1024): raise ValueError('이미지 크기를 확인하세요 (8MB 이하).')
    if data.startswith(b'\x89PNG\r\n\x1a\n'): extension='png'
    elif data.startswith(b'\xff\xd8\xff'): extension='jpg'
    elif data[:4]==b'RIFF' and data[8:12]==b'WEBP': extension='webp'
    else: raise ValueError('PNG/JPG/WebP 파일만 지원합니다.')
    image_dir=ROOT/'assets'/'generated'; image_dir.mkdir(parents=True,exist_ok=True)
    stamp=datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S_%f')
    safe_prefix=re.sub(r'[^A-Za-z0-9_-]+','_',str(prefix))[:40] or 'package'
    name=f'{safe_prefix}_{stamp}.{extension}'
    (image_dir/name).write_bytes(data)
    local_path='assets/generated/'+name
    if CLOUD_ENABLED:
        try: cloud_upload_image(local_path)
        except Exception as exc: raise ValueError('클라우드 업로드 실패: '+str(exc)) from exc
    return local_path

def save_scene_zip(body):
    data=_decode_upload_base64(body.get('zip_base64'), '이미지 ZIP')
    if len(data)>50*1024*1024: raise ValueError('이미지 ZIP은 50MB 이하만 가능합니다.')
    try: z=zipfile.ZipFile(BytesIO(data))
    except Exception as exc: raise ValueError('ZIP 파일을 열 수 없습니다.') from exc
    members=[]
    for info in z.infolist():
        if info.is_dir(): continue
        name=info.filename.replace('\\','/')
        if name.startswith('/') or '..' in name.split('/'): raise ValueError('ZIP 내부 경로가 올바르지 않습니다.')
        low=name.lower()
        if low.endswith(('.png','.jpg','.jpeg','.webp')):
            if info.file_size>8*1024*1024: raise ValueError('이미지 한 장은 8MB 이하만 가능합니다: '+Path(name).name)
            members.append((name,info))
    if not members: raise ValueError('ZIP 안에 PNG/JPG/WebP 이미지가 없습니다.')
    if len(members)>100: raise ValueError('이미지는 한 번에 100장까지 등록할 수 있습니다.')
    files=[]
    for name,info in sorted(members,key=lambda x:x[0].lower()):
        raw=z.read(info)
        image=_store_scene_bytes(raw,'zip')
        files.append({'filename':Path(name).name,'image':image})
    return {'files':files,'count':len(files)}

def import_content_package(body):
    data=_decode_upload_base64(body.get('zip_base64'), '콘텐츠 패키지 ZIP')
    if len(data)>50*1024*1024: raise ValueError('콘텐츠 패키지 ZIP은 50MB 이하만 가능합니다.')
    try: z=zipfile.ZipFile(BytesIO(data))
    except Exception as exc: raise ValueError('콘텐츠 패키지 ZIP을 열 수 없습니다.') from exc
    xlsx=[]; images=[]
    for info in z.infolist():
        if info.is_dir(): continue
        name=info.filename.replace('\\','/')
        if name.startswith('/') or '..' in name.split('/'): raise ValueError('ZIP 내부 경로가 올바르지 않습니다.')
        low=name.lower()
        if low.endswith('.xlsx'): xlsx.append((name,info))
        elif low.endswith(('.png','.jpg','.jpeg','.webp')): images.append((name,info))
    if len(xlsx)!=1: raise ValueError('ZIP 안에는 콘텐츠 엑셀(.xlsx) 파일이 정확히 1개 있어야 합니다.')
    if not images: raise ValueError('ZIP 안에 카드 이미지가 없습니다.')
    if len(images)>100: raise ValueError('이미지는 한 번에 100장까지 등록할 수 있습니다.')
    xlsx_bytes=z.read(xlsx[0][1])
    if len(xlsx_bytes)>6_000_000: raise ValueError('엑셀은 6MB 이하만 가능합니다.')
    parsed=parse_content_xlsx_preview({'file_base64':base64.b64encode(xlsx_bytes).decode('ascii')})
    valid_rows=[r for r in parsed.get('rows',[]) if r.get('valid') is not False]
    if not valid_rows: raise ValueError('엑셀에 등록 가능한 콘텐츠가 없습니다.')
    stored=[]
    for name,info in sorted(images,key=lambda x:x[0].lower()):
        if info.file_size>8*1024*1024: raise ValueError('이미지 한 장은 8MB 이하만 가능합니다: '+Path(name).name)
        image=_store_scene_bytes(z.read(info),'package')
        stored.append({'filename':Path(name).name,'image':image})
    return {'sheet':parsed.get('sheet'),'rows':valid_rows,'files':stored,'invalid_rows':parsed.get('invalid_rows',[]),'xlsx_filename':Path(xlsx[0][0]).name}

def save_uploaded_scene(body):
    """Store an uploaded candidate without changing the card's published image."""
    encoded=body.get('image_base64','')
    if not isinstance(encoded,str) or len(encoded)>12_000_000:
        raise ValueError('업로드 이미지 크기가 허용 범위를 초과했습니다.')
    try:data=base64.b64decode(encoded,validate=True)
    except Exception as exc:raise ValueError('이미지 데이터 형식이 잘못되었습니다.') from exc
    if not (100<len(data)<=8*1024*1024):raise ValueError('이미지 크기를 확인하세요 (8MB 이하).')
    if data.startswith(b'\x89PNG\r\n\x1a\n'):extension='png'
    elif data.startswith(b'\xff\xd8\xff'):extension='jpg'
    elif data[:4]==b'RIFF' and data[8:12]==b'WEBP':extension='webp'
    else:raise ValueError('PNG/JPG/WebP 파일만 지원합니다.')
    return {'image':_store_scene_bytes(data,'uploaded')}


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def end_headers(self):
        # Prevent stale JavaScript or HTML after local upgrades.
        self.send_header('Cache-Control', 'no-store, max-age=0')
        super().end_headers()

    def send_json(self, code, data):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        route=self.path.split('?',1)[0]
        if route.endswith(('.env','.pem','.key')) or route.startswith('/.git') or route.startswith('/__pycache__') or route.startswith('/logs'):
            return self.send_json(404,{'error':'찾을 수 없는 페이지입니다.'})
        if route == '/api/public-image':
            # Only published+validated scene images are public.
            from urllib.parse import parse_qs
            rel=parse_qs(urlparse(self.path).query).get('path',[''])[0]
            try:
                asset_path(rel)
                if not CLOUD_ENABLED:raise ValueError('클라우드 미설정')
                published=supa_request('/rest/v1/rpc/onecut_public_cards',method='POST',data={},service=True)
                if not isinstance(published,list) or rel not in [c.get('sceneImage') for c in published]:
                    return self.send_json(404,{'error':'공개되지 않은 이미지입니다.'})
                local=asset_path(rel)
                if local.is_file():data=local.read_bytes();mime=mimetypes.guess_type(str(local))[0] or 'image/png'
                else:data,mime=cloud_image_bytes(rel)
                self.send_response(200)
                self.send_header('Content-Type',mime)
                self.send_header('Content-Length',str(len(data)))
                self.send_header('Cache-Control','private, max-age=60')
                self.end_headers();self.wfile.write(data)
            except ValueError as exc:return self.send_json(404,{'error':str(exc)})
            return
        if route == '/api/health':
            return self.send_json(200, {'server':'onecut-v5.3.0', 'ai_key_set':bool(os.environ.get('OPENAI_API_KEY','').strip()),
                'cloud_configured':CLOUD_ENABLED,'cloud_missing':[name for name in ('SUPABASE_URL','SUPABASE_ANON_KEY','SUPABASE_SERVICE_ROLE_KEY') if not os.environ.get(name)],
                'message':'설정 여부만 표시합니다. 비밀키 값은 공개하지 않습니다.'})
        if route == '/api/cloud/status':
            return self.send_json(200, {'configured':CLOUD_ENABLED})
        if route == '/api/admin/catalog':
            try:
                require_admin(self.headers.get('Authorization','')); return self.send_json(200,admin_catalog())
            except PermissionError as err:return self.send_json(403,{'error':str(err)})
            except ValueError as err:return self.send_json(400,{'error':str(err)})
        if route == '/api/admin/backup':
            try:
                require_admin(self.headers.get('Authorization','')); return self.send_json(200,admin_backup())
            except PermissionError as err:return self.send_json(403,{'error':str(err)})
        if route == '/api/admin/report':
            try:
                require_admin(self.headers.get('Authorization','')); q=parse_qs(urlparse(self.path).query); return self.send_json(200,admin_report(q.get('target_type',[''])[0],q.get('target_id',[''])[0],q.get('days',['7'])[0]))
            except PermissionError as err:return self.send_json(403,{'error':str(err)})
            except ValueError as err:return self.send_json(400,{'error':str(err)})
        if route.startswith('/api/admin/recommendations/'):
            try:
                if not CLOUD_ENABLED:return self.send_json(503,{'error':'Supabase 설정이 없습니다.'})
                require_admin(self.headers.get('Authorization',''))
                from urllib.parse import parse_qs
                q=parse_qs(urlparse(self.path).query); count=q.get('count',['5'])[0]
                return self.send_json(200,admin_student_recommendation(route.rsplit('/',1)[-1],count))
            except PermissionError as err:return self.send_json(403,{'error':str(err)})
            except ValueError as err:return self.send_json(400,{'error':str(err)})
        if route.startswith('/api/admin/students/'):
            try:
                if not CLOUD_ENABLED:return self.send_json(503,{'error':'Supabase 설정이 없습니다.'})
                require_admin(self.headers.get('Authorization',''))
                return self.send_json(200,admin_student_detail(route.rsplit('/',1)[-1]))
            except PermissionError as err:return self.send_json(403,{'error':str(err)})
            except ValueError as err:return self.send_json(400,{'error':str(err)})
        if route == '/api/admin/groups':
            try:
                if not CLOUD_ENABLED:return self.send_json(503,{'error':'Supabase 설정이 없습니다.'})
                require_admin(self.headers.get('Authorization',''))
                return self.send_json(200,admin_groups_overview())
            except PermissionError as err:return self.send_json(403,{'error':str(err)})
            except ValueError as err:return self.send_json(400,{'error':str(err)})
        if route == '/api/student/assignments':
            try:
                token,user=student_identity(self.headers.get('Authorization',''))
                return self.send_json(200,student_assignments(user['id']))
            except PermissionError as err:return self.send_json(401,{'error':str(err)})
            except ValueError as err:return self.send_json(400,{'error':str(err)})
        if route in ('/api/cloud/cards','/api/cloud/me','/api/cloud/published','/api/admin/students'):
            try:
                if not CLOUD_ENABLED:return self.send_json(503,{'error':'Supabase 설정이 없습니다.'})
                if route=='/api/cloud/published':return self.send_json(200,cloud_published())
                adm=require_admin(self.headers.get('Authorization',''))
                if route=='/api/cloud/me':return self.send_json(200,{'user':adm})
                if route=='/api/admin/students':return self.send_json(200,admin_student_overview())
                return self.send_json(200,cloud_get_cards(adm))
            except PermissionError as err:return self.send_json(403,{'error':str(err)})
            except ValueError as err:return self.send_json(400,{'error':str(err)})
        if route in ('/api/student/me','/api/student/progress'):
            try:
                token,user=student_identity(self.headers.get('Authorization',''))
                if route=='/api/student/me':return self.send_json(200,{'user':user})
                return self.send_json(200,{'progress':student_request(token)})
            except PermissionError as err:return self.send_json(401,{'error':str(err)})
            except ValueError as err:return self.send_json(400,{'error':str(err)})
        if route == '/api/status':
            return self.send_json(200, {'ready': bool(os.environ.get('OPENAI_API_KEY','').strip()),
                                        'text_model': AI_MODEL, 'image_model': IMAGE_MODEL})
        return super().do_GET()

    def do_POST(self):
        route = self.path.split('?',1)[0]
        if route == '/api/images/check':
            # Local-only diagnostic. Never disclose arbitrary filesystem paths.
            try:
                length=int(self.headers.get('Content-Length','0'))
                if length<1 or length>65536:return self.send_json(400,{'error':'요청 크기 오류'})
                body=json.loads(self.rfile.read(length).decode('utf-8'))
                paths=body.get('paths',[])
                if not isinstance(paths,list) or len(paths)>500:return self.send_json(400,{'error':'이미지 경로 목록 오류'})
                results={}
                for entry in paths:
                    if not isinstance(entry,str) or len(entry)>600:continue
                    path=entry.replace('\\','/')
                    if path.startswith('http://') or path.startswith('https://'):
                        results[entry]={'status':'remote','message':'원격 이미지는 브라우저에서 확인'}
                    elif not path.startswith('assets/') or '..' in path.split('/') or '?' in path or '#' in path:
                        results[entry]={'status':'invalid','message':'허용되지 않은 이미지 경로'}
                    else:
                        target=(ROOT / path).resolve()
                        assets=(ROOT/'assets').resolve()
                        if not target.is_relative_to(assets):
                            results[entry]={'status':'invalid','message':'허용되지 않은 이미지 경로'}
                        else:
                            results[entry]={'status':'ok' if target.is_file() else 'missing','message':'파일 확인' if target.is_file() else '이미지 파일 없음'}
                return self.send_json(200,{'results':results})
            except Exception as err:return self.send_json(400,{'error':'이미지 검사 실패: '+str(err)})
        if route not in ('/api/student/login','/api/student/signup','/api/student/refresh','/api/student/progress','/api/student/exam-attempt','/api/student/password-reset','/api/student/update-password') and route.startswith('/api/student/'):
            return self.send_json(404,{'error':'알 수 없는 학생 API입니다.'})
        if route not in ('/api/content-xlsx-preview', '/api/content-package-upload', '/api/scene-zip-upload', '/api/anchors', '/api/mnemonics', '/api/evaluate', '/api/verify', '/api/image-review', '/api/image', '/api/research/feedback', '/api/cloud/login', '/api/cloud/refresh', '/api/cloud/cards', '/api/cloud/migrate', '/api/scene-upload', '/api/cloud/image-urls', '/api/student/login', '/api/student/signup', '/api/student/refresh', '/api/student/progress', '/api/student/exam-attempt', '/api/student/password-reset', '/api/student/update-password', '/api/admin/groups/action', '/api/admin/catalog/action', '/api/admin/curriculum/preview', '/api/admin/curriculum/apply'):
            return self.send_json(404, {'error': '알 수 없는 API 주소입니다.'})
        try:
            length = int(self.headers.get('Content-Length', '0'))
            if not 0 < length <= MAX_BODY:
                raise ValueError('요청 크기가 허용 범위를 벗어났습니다.')
            body = json.loads(self.rfile.read(length).decode('utf-8'))
            if route == '/api/admin/catalog/action':
                if not CLOUD_ENABLED: raise ValueError('Supabase 클라우드를 먼저 연결하세요.')
                actor=require_admin(self.headers.get('Authorization','')); return self.send_json(200,{'result':admin_catalog_action(actor,body)})
            if route == '/api/admin/curriculum/preview':
                if not CLOUD_ENABLED: raise ValueError('Supabase 클라우드를 먼저 연결하세요.')
                require_admin(self.headers.get('Authorization','')); return self.send_json(200,admin_curriculum_preview(body))
            if route == '/api/admin/curriculum/apply':
                if not CLOUD_ENABLED: raise ValueError('Supabase 클라우드를 먼저 연결하세요.')
                actor=require_admin(self.headers.get('Authorization','')); return self.send_json(200,admin_curriculum_apply(actor,body))
            if route == '/api/admin/groups/action':
                if not CLOUD_ENABLED: raise ValueError('Supabase 클라우드를 먼저 연결하세요.')
                actor=require_admin(self.headers.get('Authorization',''))
                return self.send_json(200,{'result':admin_groups_action(actor,body)})
            if route == '/api/student/update-password':
                if not CLOUD_ENABLED:raise ValueError('Supabase 클라우드를 먼저 연결하세요.')
                token,user=student_identity(self.headers.get('Authorization',''))
                password=str(body.get('password',''))
                if len(password)<8 or len(password)>500:raise ValueError('새 비밀번호는 8자 이상 입력해 주세요.')
                supa_request('/auth/v1/user',method='PUT',data={'password':password},headers={'Authorization':'Bearer '+token})
                return self.send_json(200,{'updated':True,'user_id':user['id']})
            if route == '/api/student/password-reset':
                if not CLOUD_ENABLED:raise ValueError('Supabase 클라우드를 먼저 연결하세요.')
                email=str(body.get('email','')).strip();
                if not email or '@' not in email or len(email)>254: raise ValueError('이메일을 확인하세요.')
                supa_request('/auth/v1/recover',method='POST',data={'email':email}); return self.send_json(200,{'sent':True})
            if route == '/api/student/exam-attempt':
                if not CLOUD_ENABLED:raise ValueError('Supabase 클라우드를 먼저 연결하세요.')
                token,user=student_identity(self.headers.get('Authorization','')); return self.send_json(200,{'result':record_exam_attempt(user['id'],body)})
            if route in ('/api/student/login','/api/student/signup','/api/student/refresh','/api/student/progress'):
                if not CLOUD_ENABLED:raise ValueError('Supabase 클라우드를 먼저 연결하세요.')
                if route=='/api/student/progress':
                    token,user=student_identity(self.headers.get('Authorization',''))
                    updated=student_request(token,'POST',body.get('progress'))
                    return self.send_json(200,{'saved':True,'user_id':user['id']})
                if route=='/api/student/refresh':
                    refresh=str(body.get('refresh_token',''))
                    if not refresh or len(refresh)>4096:raise ValueError('세션 갱신 정보 오류')
                    session=supa_request('/auth/v1/token?grant_type=refresh_token',method='POST',data={'refresh_token':refresh})
                else:
                    email=str(body.get('email','')).strip();password=str(body.get('password',''))
                    if not email or '@' not in email or len(email)>254 or not password or len(password)>500:raise ValueError('이메일 또는 비밀번호를 확인해 주세요.')
                    if route=='/api/student/signup':
                        session=supa_request('/auth/v1/signup',method='POST',data={'email':email,'password':password})
                    else:
                        session=supa_request('/auth/v1/token?grant_type=password',method='POST',data={'email':email,'password':password})
                return self.send_json(200,{'access_token':session.get('access_token'),
                    'refresh_token':session.get('refresh_token'),
                    'user':{'id':(session.get('user') or {}).get('id'), 'email':(session.get('user') or {}).get('email')},
                    'confirmation_required':not bool(session.get('access_token'))})
            if route in ('/api/cloud/login','/api/cloud/refresh'):
                if not CLOUD_ENABLED:raise ValueError('클라우드 연결이 설정되지 않았습니다.')
                if route=='/api/cloud/login':
                    email=str(body.get('email','')).strip();pw=str(body.get('password',''))
                    if not email or not pw or len(email)>300 or len(pw)>500:raise ValueError('이메일과 비밀번호를 확인하세요.')
                    session=supa_request('/auth/v1/token?grant_type=password',method='POST',data={'email':email,'password':pw})
                else:
                    refresh=str(body.get('refresh_token',''))
                    if not refresh or len(refresh)>4096:raise ValueError('세션 갱신 토큰이 없습니다.')
                    session=supa_request('/auth/v1/token?grant_type=refresh_token',method='POST',data={'refresh_token':refresh})
                admin=require_admin('Bearer '+session.get('access_token',''))
                return self.send_json(200,{'access_token':session['access_token'],'refresh_token':session.get('refresh_token',''),'user':admin})
            if CLOUD_ENABLED:
                admin=require_admin(self.headers.get('Authorization',''))
            if route == '/api/cloud/image-urls':
                if not CLOUD_ENABLED:raise ValueError('클라우드 모드가 아닙니다.')
                return self.send_json(200,cloud_admin_image_urls(body.get('paths')))
            if route == '/api/cloud/cards':
                if not CLOUD_ENABLED:raise ValueError('클라우드 모드가 아닙니다.')
                result=cloud_save_cards(body.get('cards'),body.get('expected_version'),admin)
            elif route == '/api/cloud/migrate':
                if not CLOUD_ENABLED:raise ValueError('클라우드 모드가 아닙니다.')
                result=cloud_migrate_legacy(body.get('cards'),body.get('expected_version'),admin)
            elif route=='/api/content-xlsx-preview':result=parse_content_xlsx_preview(body)
            elif route=='/api/content-package-upload':result=import_content_package(body)
            elif route=='/api/scene-zip-upload':result=save_scene_zip(body)
            elif route=='/api/anchors':result=research_phonetic_anchors(body)
            elif route=='/api/mnemonics':result=generate_mnemonics(body)
            elif route=='/api/evaluate':result=evaluate_mnemonic(body)
            elif route=='/api/verify':result=verify_mnemonic_strict(body)
            elif route=='/api/image-review':result=review_generated_image(body)
            elif route=='/api/research/feedback':result=record_research_feedback(body)
            elif route=='/api/scene-upload':result=save_uploaded_scene(body)
            else:result=generate_image(body)
            self.send_json(200, result)
        except PermissionError as exc:
            self.send_json(403, {'error': str(exc)[:450]})
        except (ValueError, KeyError, TypeError, json.JSONDecodeError) as exc:
            self.send_json(400, {'error': str(exc)[:450]})
        except Exception as exc:
            print('Server error:', type(exc).__name__, str(exc)[:200]); log_event('server_error',str(exc),route=route,error_type=type(exc).__name__)
            self.send_json(500, {'error': '서버 처리 중 오류가 발생했습니다. CMD 창의 오류 메시지를 확인해 주세요.'})


if __name__ == '__main__':
    print(f'한컷영단어 v{VERSION} http://{HOST}:{PORT}/admin.html')
    print('모드:', 'Supabase 클라우드 인증·DB 연동' if CLOUD_ENABLED else '로컬 데모 (Supabase 미설정)')
    if not os.environ.get('OPENAI_API_KEY', '').strip():
        print('참고: OPENAI_API_KEY가 없어 AI 생성 버튼은 작동하지 않습니다.')
    print('CMD 창을 종료하면 로컬 서버도 종료됩니다.')
    ThreadingHTTPServer((HOST, PORT), Handler).serve_forever()
