'use strict';
(function(){
const KEY='onecut_integrated_v2_cards', EVENT='onecut:cards-changed';
const seeds=[
{id:'idiot-n',word:'idiot',pos:'명사',ipa:'[ˈɪd.i.ət]',hangul:'이디엇',meaning:'바보, 멍청한 사람',stress:[0,2],mnemonic:'이디오 커피를 이디아로 착각해 혼자 기다리다, “아, 나 진짜 바보다!”',scene:'이디오 커피 앞에서 약속 장소를 잘못 읽은 것을 깨닫고 이마를 짚는다.',image:'assets/idiot.webp',sceneImage:'assets/preview/idiot.webp',sceneImageMode:'legacy-crop',example:'He felt like an idiot for forgetting the meeting.',translation:'그는 약속을 잊어버려 바보 같다고 느꼈다.',status:'draft',scores:[25,24,24,19],flags:[],imageChecked:false,sourceChecked:false,notes:'이디오 커피 간판과 이디아 착각이 구분되어야 함. 기존 시안은 재검수 필요.'},
{id:'immense-adj',word:'immense',pos:'형용사',ipa:'[ɪˈmens]',hangul:'이멘스',meaning:'거대한, 엄청난',stress:[2,7],mnemonic:'이 맨션, 엄청 크네!',scene:'작은 인물이 거대한 맨션을 올려다보며 놀란다.',image:'assets/immense.webp',sceneImage:'assets/preview/immense.webp',sceneImageMode:'legacy-crop',example:'The building was immense.',translation:'그 건물은 엄청나게 컸다.',status:'draft',scores:[25,24,25,19],flags:[],imageChecked:false,sourceChecked:false,notes:'기존 시안은 최종 검수 전.'},
{id:'impenetrable-adj',word:'impenetrable',pos:'형용사',ipa:'[ɪmˈpen.ɪ.trə.bəl]',hangul:'임페너트러블',meaning:'뚫을 수 없는, 통과할 수 없는',stress:[2,5],mnemonic:'이 펜이 들어? 불가능해!',scene:'두꺼운 철문을 펜으로 뚫으려다 실패한다.',image:'assets/impenetrable.webp',sceneImage:'assets/preview/impenetrable.webp',sceneImageMode:'legacy-crop',example:'The forest seemed impenetrable.',translation:'그 숲은 통과할 수 없을 것처럼 보였다.',status:'draft',scores:[13,16,20,18],flags:['awkward'],imageChecked:false,sourceChecked:false,notes:'억지 연상으로 보류. 원칙상 공개 금지.'},
{id:'rebel-n',word:'rebel',pos:'명사',ipa:'[ˈreb.əl]',hangul:'레벌',meaning:'반항자, 반역자',stress:[0,3],mnemonic:'내가 벌 받으라고? 절대 안 해!',scene:'벌을 받으라는 지시에 팔짱을 끼고 반발하는 인물.',image:'assets/rebel-n.webp',sceneImage:'assets/preview/rebel-n.webp',sceneImageMode:'legacy-crop',example:'The rebel refused to obey the king.',translation:'그 반역자는 왕에게 복종하기를 거부했다.',status:'draft',scores:[21,23,24,19],flags:[],imageChecked:false,sourceChecked:false,notes:'명사 강세는 첫 음절. 연결 자연스러움 검수.'},
{id:'rebel-v',word:'rebel',pos:'동사',ipa:'[rɪˈbel]',hangul:'리벨',meaning:'반항하다, 저항하다',stress:[2,5],mnemonic:'네가 벨 눌러! 싫어요!',scene:'벨을 누르라는 지시를 들은 사람이 거절하며 반항한다.',image:'assets/rebel-v.webp',sceneImage:'assets/preview/rebel-v.webp',sceneImageMode:'legacy-crop',example:'The students rebelled against the rule.',translation:'학생들은 그 규칙에 반항했다.',status:'draft',scores:[22,23,24,19],flags:[],imageChecked:false,sourceChecked:false,notes:'동사는 둘째 음절 강세.'},
{id:'allow-v',word:'allow',pos:'동사',ipa:'[əˈlaʊ]',hangul:'얼라우',meaning:'허락하다',stress:[2,5],mnemonic:'',scene:'',image:'',example:'',translation:'',status:'rejected',scores:[0,0,0,0],flags:['awkward'],imageChecked:false,sourceChecked:false,notes:'기존 억지 연상 폐기'},
{id:'belong-v',word:'belong',pos:'동사',ipa:'[bɪˈlɔːŋ]',hangul:'빌롱',meaning:'~에 속하다',stress:[2,6],mnemonic:'',scene:'',image:'',example:'',translation:'',status:'rejected',scores:[0,0,0,0],flags:['awkward'],imageChecked:false,sourceChecked:false,notes:'기존 억지 연상 폐기'}
];
const clone=x=>JSON.parse(JSON.stringify(x));

function escapeHtml(v){return String(v??'').replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]))}

// v5.2.5: 강세는 추측하지 않고 데이터의 [강세] 표기를 최우선으로 사용한다.
// 예: con[tam]inate, contami[na]tion, [core]
const legacyStressMarks={
 core:'[core]', contaminate:'con[tam]inate', contamination:'contami[na]tion',
 torture:'[tor]ture', germ:'[germ]', measure:'[mea]sure', measurable:'[mea]surable',
 unmeasurable:'un[mea]surable', threat:'[threat]', threaten:'[threat]en', spoil:'[spoil]'
};
function parseStressMark(mark,word){
 const m=String(mark||'').trim(), w=String(word||'').trim();
 const hit=m.match(/^([^\[\]]*)\[([^\[\]]+)\]([^\[\]]*)$/);
 if(!hit)return null;
 const plain=hit[1]+hit[2]+hit[3];
 if(plain.toLowerCase()!==w.toLowerCase())return null;
 return [hit[1].length,hit[1].length+hit[2].length];
}
function stressMarkFor(c){
 const explicit=String(c?.stressMark||c?.excelImport?.stressMark||'').trim();
 if(explicit)return explicit;
 // 기존 17개 엑셀 카드만 1회 호환 보정. 신규 데이터는 반드시 엑셀 강세 표시를 사용한다.
 if(c?.excelImport){const legacy=legacyStressMarks[String(c?.word||'').toLowerCase()];if(legacy)return legacy;}
 return '';
}
function formatStressMark(word,stress){
 const w=String(word||'');
 if(!Array.isArray(stress)||stress.length!==2)return w;
 const a=Number(stress[0]),b=Number(stress[1]);
 if(!Number.isInteger(a)||!Number.isInteger(b)||a<0||b<=a||b>w.length)return w;
 return w.slice(0,a)+'['+w.slice(a,b)+']'+w.slice(b);
}
function stressFor(c){
 const w=String(c?.word||''), marked=parseStressMark(stressMarkFor(c),w);
 if(marked)return marked;
 const raw=Array.isArray(c?.stress)?c.stress:[];
 const a=Number(raw[0]),b=Number(raw[1]);
 if(Number.isInteger(a)&&Number.isInteger(b)&&a>=0&&b>a&&b<=w.length)return [a,b];
 return [0,0];
}
function wordMarkup(c){
 const w=String(c.word||''),st=stressFor(c),start=Math.max(0,Number(st?.[0]||0)),end=Math.min(w.length,Number(st?.[1]||0));
 if(end<=start)return escapeHtml(w);
 return escapeHtml(w.slice(0,start))+'<span class="stress">'+escapeHtml(w.slice(start,end))+'</span>'+escapeHtml(w.slice(end));
}
function posFor(word,meaning){
 const w=String(word||'').toLowerCase(), m=String(meaning||'');
 if(w==='core'||w==='contamination'||w==='germ'||w==='threat')return '명사';
 if(w==='measurable'||w==='unmeasurable')return '형용사';
 if(w==='measure'&&m.includes('조치'))return '명사';
 if(['contaminate','threaten','spoil'].includes(w))return '동사';
 if(w==='torture')return m.includes('고문')?'명사':'동사';
 if(w==='measure')return '동사';
 return '미확인';
}
function sceneSrc(c){return c.sceneImage||((c.sceneImageMode==='clean')?c.image:'');}
function cardMarkup(c,opts={}){
 const esc=escapeHtml, img=sceneSrc(c),legacy=c.sceneImageMode!=='clean';
 return `<section class="onecut-card" aria-label="${esc(c.word)} 발음연상 카드">
   <div class="onecut-word">${wordMarkup(c)}</div>
   <div class="onecut-phon">${esc(c.ipa||'')} <span class="onecut-hangul">${esc(c.hangul||'')}</span></div>
   <div class="onecut-meaning">${esc(c.meaning||'')}</div>
   <div class="onecut-scene ${legacy?'legacy':'clean'}">${img?`<img src="${esc(img)}" alt="${esc(c.scene||'연상 그림')}" loading="lazy" onerror="this.hidden=true;this.parentElement.classList.add('scene-image-missing');let note=this.parentElement.querySelector('.image-missing-notice');if(!note){note=document.createElement('div');note.className='image-missing-notice';note.textContent='이미지를 불러오지 못했습니다. 관리자에서 이미지 파일을 다시 확인하세요.';this.parentElement.append(note)}">`:'<div class="onecut-placeholder">한 컷 장면 이미지가 아직 없습니다.<small>장면 전용 그림을 생성하거나 등록해 주세요.</small></div>'}</div>
   <div class="onecut-mnemonic">${esc(c.mnemonic||'')}</div>
  </section>${legacy&&img?'<p class="preview-note">기존 포스터를 잘라 만든 임시 장면 미리보기입니다. 남아 있는 글자가 있을 수 있어 최종 승인 전 장면 전용 그림으로 교체하세요.</p>':''}`;
}
function audit(c){
 const problems=[];
 if(!c.word?.trim()||!c.pos?.trim()||!c.ipa?.trim()||!c.meaning?.trim())problems.push('단어·품사·IPA·뜻 필수');
 if(!c.mnemonic?.trim()||!c.scene?.trim())problems.push('연상문장과 한 컷 장면 미입력');
 if(!c.example?.trim()||!c.translation?.trim())problems.push('예문·해석 미입력');
 if(!Array.isArray(c.stress)||c.stress.length!==2||!Number.isInteger(c.stress[0])||!Number.isInteger(c.stress[1])||c.stress[0]<0||c.stress[1]>String(c.word||'').length||c.stress[0]>=c.stress[1])problems.push('강세 위치 불일치');
 if(!c.evaluation||!['ai','manual'].includes(c.evaluation.status))problems.push('AI 품질평가 또는 사유 있는 관리자 수동평가 필요');
 if((c.scores||[]).reduce((a,n)=>a+Number(n||0),0)<85)problems.push('품질평가 85점 미만');
 if(c.flags?.length)problems.push('반려 플래그 있음');
 if(c.sceneImageMode!=='clean'||!c.sceneImage?.trim())problems.push('검수 완료된 장면 전용 그림 없음');
 if(!c.sourceChecked)problems.push('발음·연상 사람 검수 필요');
 if(!c.imageChecked)problems.push('이미지 사람 검수 필요');
 if(c.speechRequired&&!c.sceneSpeech?.trim())problems.push('필수 말풍선 대사 미입력');
 if(c.speechRequired&&!c.speechChecked)problems.push('그림 속 핵심 말풍선 확인 필요');
 return problems;
}
const store={wordMarkup,cardMarkup,sceneSrc,audit,stressFor,parseStressMark,stressMarkFor,formatStressMark,posFor,key:KEY,seed:()=>clone(seeds),load(){try{const a=JSON.parse(localStorage.getItem(KEY));return Array.isArray(a)?a:clone(seeds)}catch{return clone(seeds)}},save(cards){localStorage.setItem(KEY,JSON.stringify(cards));window.dispatchEvent(new Event(EVENT));},published(){return this.load().filter(c=>c.status==='published'&&this.eligible(c))},score(c){return (c.scores||[]).reduce((a,n)=>a+Number(n||0),0)},eligible(c){return !!(c.sceneImage?.trim()&&c.sceneImageMode==='clean'&&c.word?.trim()&&c.pos?.trim()&&c.meaning?.trim()&&c.ipa?.trim()&&c.mnemonic?.trim()&&c.scene?.trim())&&Array.isArray(c.stress)&&c.stress.length===2&&Number.isInteger(c.stress[0])&&Number.isInteger(c.stress[1])&&c.stress[0]>=0&&c.stress[1]>c.stress[0]&&c.stress[1]<=c.word.length},escape(v){return String(v??'').replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]))},EVENT};
window.OneCut=store;
})();
