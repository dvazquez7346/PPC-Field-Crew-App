/* ============================================================================
   SHARED APP CORE — auth (MSAL) + SharePoint-via-Graph file storage
   ============================================================================
   Used by BOTH index.html (Field Crew Directory) and drawings.html (markup
   tool). Keeping this in one file means:
     - one login for both pages (MSAL silently reuses the session)
     - one place to fix a bug or rotate a client ID
     - each page still loads/saves its OWN json file, so the crew directory
       never gets slower just because the drawing tool's data grows

   HOW TO USE IN A NEW PAGE:
     1. Load the MSAL browser SDK via CDN (see bottom of this file for the
        exact <script> tag) BEFORE this file.
     2. Load this file.
     3. Call AppCore.initMsal() once, then AppCore.signIn()/getGraphToken().
     4. Call AppCore.readJsonFile('your-file-name.json', defaultObj) and
        AppCore.writeJsonFile('your-file-name.json', dataObj, lastModified)
        to read/write YOUR OWN data file — separate from the crew directory's.
   ========================================================================= */

const AppCore = (function(){

  /* ---- Same tenant/site config as the existing crew directory app ---- */
  const MSAL_CLIENT_ID = "ccb32922-c056-4605-96e9-c4f284944cb2";
  const MSAL_TENANT_ID = "49bfc544-45fd-4913-bd81-5503c930b3d8";
  const SITE_HOSTNAME  = "1ppc0.sharepoint.com";
  const SITE_PATH      = "/sites/m365appbuilder-field-crew-directory-8280";
  const GRAPH_SCOPES   = ["Sites.ReadWrite.All"];

  let msalInstance = null;
  let msalLoadError = null;
  let cachedSiteId = null;

  /* ============================= AUTH (MSAL) ============================= */

  function initMsal(){
    if(msalInstance) return msalInstance;
    try{
      if(typeof msal === 'undefined'){
        throw new Error('MSAL library not loaded — add the CDN <script> tag before shared-app-core.js');
      }
      msalInstance = new msal.PublicClientApplication({
        auth: {
          clientId: MSAL_CLIENT_ID,
          authority: `https://login.microsoftonline.com/${MSAL_TENANT_ID}`,
          redirectUri: window.location.origin + window.location.pathname,
        },
        cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false },
      });
    }catch(e){
      console.error('MSAL setup failed', e);
      msalLoadError = e.message || String(e);
    }
    return msalInstance;
  }

  function isMobileDevice(){
    return /Android|iPhone|iPad|iPod|Mobile/i.test(navigator.userAgent) || window.matchMedia('(max-width: 899px)').matches;
  }
  function isInIframe(){
    try{ return window.parent !== window; }catch(e){ return true; }
  }

  // Returns the signed-in account, or null. Since cacheLocation is localStorage
  // and both pages use the same clientId/authority, signing in on one page means
  // the other page sees the same account with zero extra login prompts.
  function getActiveAccount(){
    if(!msalInstance) return null;
    const accounts = msalInstance.getAllAccounts();
    return accounts.length ? accounts[0] : null;
  }

  async function getGraphToken(){
    const account = getActiveAccount();
    if(!account) throw new Error('Not signed in');
    const req = { scopes: GRAPH_SCOPES, account };
    try{
      const res = await msalInstance.acquireTokenSilent(req);
      return res.accessToken;
    }catch(e){
      if(isMobileDevice() && !isInIframe()){
        await msalInstance.acquireTokenRedirect(req);
        return; // page navigates away; caller's page should call handleRedirectPromise on load
      }
      const res = await msalInstance.acquireTokenPopup(req);
      return res.accessToken;
    }
  }

  // onSuccess(account) / onError(message) let each page hook its own state+render
  // without this shared file knowing anything about either page's UI.
  async function signIn(onSuccess, onError){
    if(isMobileDevice() && !isInIframe()){
      try{
        await msalInstance.loginRedirect({ scopes: GRAPH_SCOPES });
      }catch(e){
        console.error('Sign-in redirect failed', e);
        if(onError) onError((e && e.errorMessage) || (e && e.message) || 'Sign-in failed');
      }
      return; // page navigates away; execution resumes via handleRedirectPromise on reload
    }
    try{
      const res = await msalInstance.loginPopup({ scopes: GRAPH_SCOPES });
      if(onSuccess) onSuccess(res.account);
    }catch(e){
      console.error('Sign-in failed', e);
      if(onError) onError((e && e.errorMessage) || (e && e.message) || 'Sign-in failed');
    }
  }

  function signOut(){
    const account = getActiveAccount();
    if(account && msalInstance){
      msalInstance.logoutPopup({ account }).catch(()=>{});
    }
  }

  // Call once on page load (after initMsal) to finish a redirect-based sign-in
  // on mobile. Resolves to the account if a redirect just completed, else null.
  async function handleRedirectPromise(){
    if(!msalInstance) return null;
    try{
      const res = await msalInstance.handleRedirectPromise();
      return res ? res.account : null;
    }catch(e){
      console.error('Redirect handling failed', e);
      return null;
    }
  }

  /* ==================== STORAGE (SharePoint via Microsoft Graph) ====================
     Generalized version of the crew app's read/write pattern. Same 404-creates-empty-
     file behavior, same optimistic-concurrency check via lastModifiedDateTime — just
     parameterized by filename so each page owns its own file. This is the key to
     keeping things fast: the crew directory's file never grows because of drawing
     markup data, and vice versa. */

  async function resolveSiteId(token){
    if(cachedSiteId) return cachedSiteId;
    const res = await fetch(`https://graph.microsoft.com/v1.0/sites/${SITE_HOSTNAME}:${SITE_PATH}`, {
      headers: { Authorization: `Bearer ${token}` }
    });
    if(!res.ok) throw new Error(`Could not find SharePoint site (${res.status})`);
    const json = await res.json();
    cachedSiteId = json.id;
    return cachedSiteId;
  }

  function fileUrl(siteId, filename, suffix=''){
    return `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodeURIComponent(filename)}${suffix}`;
  }

  // Reads `filename` from the SharePoint drive. If it doesn't exist yet, creates
  // it with defaultObj (proving write access in the process) and returns that.
  // Returns { data, lastModified, canWrite, itemId }.
  async function readJsonFile(filename, defaultObj){
    const token = await getGraphToken();
    const siteId = await resolveSiteId(token);
    const metaRes = await fetch(fileUrl(siteId, filename), { headers: { Authorization: `Bearer ${token}` } });

    if(metaRes.status === 404){
      const created = await writeJsonFile(filename, defaultObj, null);
      return { data: defaultObj, lastModified: created.lastModified, canWrite: true, itemId: created.itemId };
    }
    if(metaRes.status === 403){
      throw new Error('Your Microsoft account doesn\'t have access to this SharePoint site. Ask your site admin to add you.');
    }
    if(!metaRes.ok){
      throw new Error(`Could not read ${filename} (${metaRes.status})`);
    }
    const meta = await metaRes.json();
    const contentRes = await fetch(fileUrl(siteId, filename, ':/content'), { headers: { Authorization: `Bearer ${token}` } });
    const data = await contentRes.json();
    const canWrite = await checkWriteAccess(token, siteId, meta.id, meta.description);
    return { data, lastModified: meta.lastModifiedDateTime, canWrite, itemId: meta.id };
  }

  // Probes write access with a harmless no-op PATCH (re-setting description to
  // its current value). Returns false only on a definitive 403; assumes true on
  // ambiguous errors so a flaky network doesn't wrongly lock someone out.
  async function checkWriteAccess(token, siteId, itemId, currentDescription){
    try{
      const res = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${itemId}`, {
        method: 'PATCH',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ description: currentDescription || '' }),
      });
      if(res.status === 403) return false;
      return true;
    }catch(e){
      console.error('Write-access check failed', e);
      return true;
    }
  }

  // Writes dataObj to `filename`. If lastKnownModified is provided, first checks
  // whether the file changed since the caller last read it — throws a conflict
  // error rather than silently clobbering someone else's concurrent edit.
  // Returns { lastModified, itemId }.
  async function writeJsonFile(filename, dataObj, lastKnownModified){
    const token = await getGraphToken();
    const siteId = await resolveSiteId(token);

    if(lastKnownModified){
      const metaRes = await fetch(fileUrl(siteId, filename), { headers: { Authorization: `Bearer ${token}` } });
      if(metaRes.ok){
        const meta = await metaRes.json();
        if(meta.lastModifiedDateTime && meta.lastModifiedDateTime !== lastKnownModified){
          throw Object.assign(new Error('CONFLICT'), { isConflict: true });
        }
      }
    }

    const res = await fetch(fileUrl(siteId, filename, ':/content'), {
      method: 'PUT',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(dataObj),
    });
    if(res.status === 403){
      throw Object.assign(new Error('You don\'t have permission to save changes.'), { isForbidden: true });
    }
    if(!res.ok) throw new Error(`Save failed (${res.status})`);
    const json = await res.json();
    return { lastModified: json.lastModifiedDateTime, itemId: json.id };
  }

  /* ---- Uploads a binary file (e.g. a PDF drawing) to the SAME site's drive,
     under a given folder path, for the drawing tool to use later. Kept here so
     any future page can drop files into SharePoint without re-deriving siteId
     logic. Returns the created item's id + webUrl. ---- */
  async function uploadFile(folderPath, filename, blobOrArrayBuffer, contentType){
    const token = await getGraphToken();
    const siteId = await resolveSiteId(token);
    const path = `${folderPath.replace(/\/$/, '')}/${filename}`;
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodeURIComponent(path)}:/content`,
      {
        method: 'PUT',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': contentType || 'application/octet-stream' },
        body: blobOrArrayBuffer,
      }
    );
    if(!res.ok) throw new Error(`Upload failed (${res.status})`);
    return await res.json();
  }

  return {
    initMsal, getActiveAccount, getGraphToken, signIn, signOut, handleRedirectPromise,
    readJsonFile, writeJsonFile, uploadFile,
    isMobileDevice, isInIframe,
  };
})();

/* ============================================================================
   To use this in drawings.html (or any new page), put this in <head>, in order:

   <script src="https://cdn.jsdelivr.net/npm/@azure/msal-browser@2/lib/msal-browser.min.js"></script>
   <script src="shared-app-core.js"></script>

   Then in your page's own script:

   AppCore.initMsal();
   await AppCore.handleRedirectPromise();   // finishes mobile redirect sign-in, if any
   if (AppCore.getActiveAccount()) {
     const { data, lastModified, canWrite } =
       await AppCore.readJsonFile('job-4521-markup.json', { sheets: [] });
     // ...render your drawing tool using `data`...
   } else {
     AppCore.signIn(
       (account) => { /* re-run the load above */ },
       (errMsg)  => { /* show errMsg to the user */ }
     );
   }

   To save:
   await AppCore.writeJsonFile('job-4521-markup.json', updatedData, lastModified);
   ========================================================================= */
