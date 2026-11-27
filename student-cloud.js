/* OneCut Student Cloud: credentials travel only to local server over same-origin; never exposes service key. */
'use strict';
window.OneCutStudent=(function(){
 const TOKEN='onecut_student_access',REFRESH='onecut_student_refresh';
 let access=sessionStorage.getItem(TOKEN)||'',refresh=sessionStorage.getItem(REFRESH)||'';
 async function call(path,data,authenticate=true){
   const headers={'Content-Type':'application/json'};
   if(authenticate&&access)headers.Authorization='Bearer '+access;
   const res=await fetch('/api/student/'+path,{method:data===undefined?'GET':'POST',headers,
     body:data===undefined?undefined:JSON.stringify(data),cache:'no-store'});
   let body={};try{body=await res.json()}catch{}
   if(!res.ok)throw Error(body.error||'서버 오류 '+res.status);
   return body;
 }
 function session(s){if(s.access_token){access=s.access_token;refresh=s.refresh_token||refresh;
   sessionStorage.setItem(TOKEN,access);sessionStorage.setItem(REFRESH,refresh);}return s;}
 return {
  async login(email,password){return session(await call('login',{email,password},false))},
  async signup(email,password){return session(await call('signup',{email,password},false))},
  async resume(){if(!access)return null;try{return await call('me')}catch(e){
    if(!refresh)throw e;try{session(await call('refresh',{refresh_token:refresh},false));return await call('me')}
    catch{this.logout();throw e}
  }},
  async load(){return (await call('progress')).progress||{}},
  async save(progress){return await call('progress',{progress})},
  async assignments(){return (await call('assignments')).assignments||[]},
  async examAttempt(payload){return await call('exam-attempt',payload)},
  async passwordReset(email){return await call('password-reset',{email},false)},
  async updatePassword(recoveryToken,password){const res=await fetch('/api/student/update-password',{method:'POST',headers:{'Content-Type':'application/json','Authorization':'Bearer '+recoveryToken},body:JSON.stringify({password}),cache:'no-store'});let body={};try{body=await res.json()}catch{}if(!res.ok)throw Error(body.error||'비밀번호 변경 실패');return body},
  logout(){access='';refresh='';sessionStorage.removeItem(TOKEN);sessionStorage.removeItem(REFRESH)},
  get loggedIn(){return !!access}
 };
})();
