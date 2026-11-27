/* 한컷영단어 v2.6 클라우드 브리지. 모든 DB 쓰기는 서버 권한 검증 후 수행. */
'use strict';
window.OneCutCloud=(function(){
 let enabled=false, ready=false, token=sessionStorage.getItem('onecut_admin_token')||'', refreshToken=sessionStorage.getItem('onecut_admin_refresh')||'';
 let user=null, lastError='', version=null;
 const KEY='onecut_integrated_v2_cards';
 async function request(url, options={}){
   const headers={...options.headers};if(token)headers.Authorization='Bearer '+token;
   if(options.body && !headers['Content-Type'])headers['Content-Type']='application/json';
   let res=await fetch(url,{...options,headers,cache:'no-store'});
   if(res.status===403 && token && refreshToken && !url.includes('/cloud/login') && !url.includes('/cloud/refresh')){
     try{const rr=await fetch('/api/cloud/refresh',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({refresh_token:refreshToken}),cache:'no-store'});
       if(rr.ok){const session=await rr.json();setSession(session);headers.Authorization='Bearer '+token;res=await fetch(url,{...options,headers,cache:'no-store'})}
     }catch(e){console.warn('세션 갱신 실패',e)}
   }
   let data={};
   try{data=await res.json()}catch{throw Error('서버가 JSON 응답을 반환하지 않았습니다.');}
   if(!res.ok)throw Error(data.error||`서버 오류 ${res.status}`);
   return data;
 }
 async function status(){const s=await request('/api/cloud/status');enabled=!!s.configured;return s;}
 function setSession(s){token=s.access_token;refreshToken=s.refresh_token||refreshToken;sessionStorage.setItem('onecut_admin_token',token);sessionStorage.setItem('onecut_admin_refresh',refreshToken);user=s.user;}
 async function login(email,password){const r=await request('/api/cloud/login',{method:'POST',body:JSON.stringify({email,password})});setSession(r);return r;}
 function logout(){token='';refreshToken='';user=null;sessionStorage.removeItem('onecut_admin_token');sessionStorage.removeItem('onecut_admin_refresh');}
 async function me(){const data=await request('/api/cloud/me');user=data.user;return user;}
 async function loadAdmin(){const d=await request('/api/cloud/cards');version=d.version;return d.cards;}
 async function save(cards){const d=await request('/api/cloud/cards',{method:'POST',body:JSON.stringify({cards,expected_version:version})});version=d.version;return d;}
 async function migrate(cards){const d=await request('/api/cloud/migrate',{method:'POST',body:JSON.stringify({cards,expected_version:version})});version=d.version;return d;}
 async function published(){const d=await request('/api/cloud/published');return d.cards;}
  return {status,login,logout,me,loadAdmin,save,migrate,published,request,get enabled(){return enabled},get ready(){return ready},set ready(v){ready=v},get token(){return token},get user(){return user},get lastError(){return lastError},set lastError(v){lastError=v}};
})();
