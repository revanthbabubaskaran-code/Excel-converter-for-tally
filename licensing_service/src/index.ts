interface Env {
  LICENSE_DB: D1Database;
  LICENSE_ADMIN_TOKEN: string;
  LICENSE_SIGNING_KEY: string;
}

interface LicenseRequest {
  id: string;
  device_id: string;
  request_code: string;
  status: "pending" | "approved" | "rejected";
  license_token: string | null;
  created_at: string;
  updated_at: string;
}

const PRODUCT_ID = "excel-for-tally-v1";
const ADMIN_PAGE = `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Excel for Tally — License approvals</title>
<style>
body{font:16px Segoe UI,Arial,sans-serif;max-width:1000px;margin:36px auto;padding:0 18px;color:#23313a;background:#f3f6f8}
h1{color:#123d54}.card{background:white;border:1px solid #d7e0e5;border-radius:10px;padding:18px;margin:18px 0}
input,button{font:inherit;padding:9px 12px;border:1px solid #bdcbd2;border-radius:6px}input{min-width:320px;max-width:100%}
button{cursor:pointer;background:#edf4f7}button.approve{background:#d9f0e2}button.reject{background:#f9e1e1}
.row{display:flex;gap:8px;align-items:center;flex-wrap:wrap}.request{border-top:1px solid #e1e8ec;padding:14px 0}.code{font-family:Consolas,monospace;overflow-wrap:anywhere}
#status{color:#52636d;white-space:pre-wrap}.error{color:#b3261e}small{color:#52636d}
</style></head><body>
<h1>Excel for Tally license approvals</h1>
<section class="card"><label for="adminToken">Owner access token</label><div class="row" style="margin-top:8px">
<input id="adminToken" type="password" autocomplete="off" placeholder="Paste the service admin token">
<button id="load">Load pending requests</button></div><p id="status">The token stays in this browser tab and is sent only to this service.</p></section>
<main id="requests" class="card"><small>Pending device requests appear here after you load them.</small></main>
<script>
const token=document.getElementById('adminToken'), status=document.getElementById('status'), box=document.getElementById('requests');
async function api(path,options={}){const response=await fetch(path,{...options,headers:{'Authorization':'Bearer '+token.value,'Content-Type':'application/json',...(options.headers||{})}});const data=await response.json().catch(()=>({error:'Invalid server response'}));if(!response.ok)throw new Error(data.error||'Request failed');return data}
function showStatus(text,isError=false){status.textContent=text;status.className=isError?'error':''}
async function load(){try{const data=await api('/api/admin/requests');box.replaceChildren();if(!data.requests.length){box.textContent='No pending license requests.';return}for(const item of data.requests){const row=document.createElement('section');row.className='request';const id=document.createElement('div');id.textContent='Request: '+item.id;const code=document.createElement('div');code.className='code';code.textContent=item.request_code;const created=document.createElement('small');created.textContent='Received: '+item.created_at;const actions=document.createElement('div');actions.className='row';for(const [action,label,cls] of [['approve','Approve license','approve'],['reject','Reject','reject']]){const button=document.createElement('button');button.textContent=label;button.className=cls;button.onclick=()=>decide(item.id,action);actions.append(button)}row.append(id,code,created,actions);box.append(row)}showStatus('Loaded '+data.requests.length+' pending request(s).')}catch(error){showStatus(error.message,true)}}
async function decide(id,action){try{await api('/api/admin/requests/'+encodeURIComponent(id)+'/'+action,{method:'POST',body:'{}'});showStatus(action==='approve'?'License approved and issued. The device can now check its request.':'Request rejected.');await load()}catch(error){showStatus(error.message,true)}}
document.getElementById('load').onclick=load;
</script></body></html>`;

function json(data: unknown, status = 200): Response {
  return new Response(JSON.stringify(data), { status, headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" } });
}

function errorResponse(message: string, status: number): Response {
  return json({ error: message }, status);
}

function timingSafeEqual(a: string, b: string): boolean {
  const left = new TextEncoder().encode(a);
  const right = new TextEncoder().encode(b);
  let difference = left.length ^ right.length;
  const length = Math.max(left.length, right.length);
  for (let i = 0; i < length; i += 1) difference |= (left[i] ?? 0) ^ (right[i] ?? 0);
  return difference === 0;
}

function authorizeAdmin(request: Request, env: Env): boolean {
  const authorization = request.headers.get("authorization") ?? "";
  const supplied = authorization.startsWith("Bearer ") ? authorization.slice(7) : "";
  return Boolean(env.LICENSE_ADMIN_TOKEN) && timingSafeEqual(supplied, env.LICENSE_ADMIN_TOKEN);
}

function base64Url(bytes: Uint8Array): string {
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replaceAll("+", "-").replaceAll("/", "_").replace(/=+$/, "");
}

function base64Bytes(value: string): Uint8Array {
  const binary = atob(value);
  return Uint8Array.from(binary, (character) => character.charCodeAt(0));
}

async function signLicense(deviceId: string, env: Env): Promise<string> {
  if (!env.LICENSE_SIGNING_KEY) throw new Error("License signing is not configured.");
  const privateBytes = base64Bytes(env.LICENSE_SIGNING_KEY.trim());
  const privateKey = await crypto.subtle.importKey("raw", privateBytes, { name: "Ed25519" }, false, ["sign"]);
  const payload = JSON.stringify({
    device_id: deviceId,
    issued_utc: new Date().toISOString(),
    license_id: crypto.randomUUID().replaceAll("-", "").slice(0, 16).toUpperCase(),
    product: PRODUCT_ID,
  });
  const signature = await crypto.subtle.sign("Ed25519", privateKey, new TextEncoder().encode(payload));
  return `EFT1.${base64Url(new TextEncoder().encode(payload))}.${base64Url(new Uint8Array(signature))}`;
}

async function parseRequest(request: Request): Promise<{ requestCode: string; deviceId: string } | Response> {
  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return errorResponse("Request body must be JSON.", 400);
  }
  const requestCode = typeof body === "object" && body !== null && "request_code" in body
    ? String((body as { request_code: unknown }).request_code).trim().toUpperCase()
    : "";
  if (!/^EFT1-[0-9A-F]{32}$/.test(requestCode)) return errorResponse("Device request code is invalid.", 400);
  return { requestCode, deviceId: requestCode.slice(5) };
}

async function getRequestByDevice(env: Env, deviceId: string): Promise<LicenseRequest | null> {
  return env.LICENSE_DB.prepare("SELECT * FROM license_requests WHERE device_id = ?1").bind(deviceId).first<LicenseRequest>();
}

async function submitRequest(request: Request, env: Env): Promise<Response> {
  const parsed = await parseRequest(request);
  if (parsed instanceof Response) return parsed;
  const existing = await getRequestByDevice(env, parsed.deviceId);
  const now = new Date().toISOString();
  if (!existing) {
    await env.LICENSE_DB.prepare(
      "INSERT INTO license_requests (id, device_id, request_code, status, license_token, created_at, updated_at) VALUES (?1, ?2, ?3, 'pending', NULL, ?4, ?4)",
    ).bind(crypto.randomUUID(), parsed.deviceId, parsed.requestCode, now).run();
  } else if (existing.status === "rejected") {
    await env.LICENSE_DB.prepare(
      "UPDATE license_requests SET request_code = ?1, status = 'pending', license_token = NULL, updated_at = ?2 WHERE device_id = ?3",
    ).bind(parsed.requestCode, now, parsed.deviceId).run();
  }
  const current = await getRequestByDevice(env, parsed.deviceId);
  if (!current) return errorResponse("Could not create the license request.", 500);
  return json({ request_id: current.id, status: current.status, license_token: current.status === "approved" ? current.license_token : undefined });
}

async function checkRequest(url: URL, env: Env): Promise<Response> {
  const requestCode = decodeURIComponent(url.pathname.slice("/api/v1/license-requests/".length)).trim().toUpperCase();
  if (!/^EFT1-[0-9A-F]{32}$/.test(requestCode)) return errorResponse("Device request code is invalid.", 400);
  const record = await getRequestByDevice(env, requestCode.slice(5));
  if (!record || record.request_code !== requestCode) return json({ status: "not_found" });
  return json({ status: record.status, license_token: record.status === "approved" ? record.license_token : undefined });
}

async function adminRequests(request: Request, env: Env): Promise<Response> {
  if (!authorizeAdmin(request, env)) return errorResponse("Owner access token is incorrect.", 401);
  const result = await env.LICENSE_DB.prepare(
    "SELECT id, request_code, created_at FROM license_requests WHERE status = 'pending' ORDER BY created_at ASC LIMIT 100",
  ).all<Pick<LicenseRequest, "id" | "request_code" | "created_at">>();
  return json({ requests: result.results ?? [] });
}

async function decideRequest(request: Request, env: Env, requestId: string, action: "approve" | "reject"): Promise<Response> {
  if (!authorizeAdmin(request, env)) return errorResponse("Owner access token is incorrect.", 401);
  const current = await env.LICENSE_DB.prepare("SELECT * FROM license_requests WHERE id = ?1").bind(requestId).first<LicenseRequest>();
  if (!current) return errorResponse("License request not found.", 404);
  if (current.status !== "pending") return errorResponse("This request has already been decided.", 409);
  const now = new Date().toISOString();
  if (action === "reject") {
    await env.LICENSE_DB.prepare("UPDATE license_requests SET status = 'rejected', updated_at = ?1 WHERE id = ?2").bind(now, requestId).run();
    return json({ status: "rejected" });
  }
  let token: string;
  try {
    token = await signLicense(current.device_id, env);
  } catch {
    return errorResponse("Could not issue the license. Check the service signing-key secret.", 500);
  }
  await env.LICENSE_DB.prepare(
    "UPDATE license_requests SET status = 'approved', license_token = ?1, updated_at = ?2 WHERE id = ?3",
  ).bind(token, now, requestId).run();
  return json({ status: "approved" });
}

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    const url = new URL(request.url);
    try {
      if (request.method === "GET" && url.pathname === "/health") return json({ status: "ok" });
      if (request.method === "GET" && url.pathname === "/admin") {
        return new Response(ADMIN_PAGE, { headers: { "content-type": "text/html; charset=utf-8", "cache-control": "no-store" } });
      }
      if (request.method === "POST" && url.pathname === "/api/v1/license-requests") return submitRequest(request, env);
      if (request.method === "GET" && url.pathname.startsWith("/api/v1/license-requests/")) return checkRequest(url, env);
      if (request.method === "GET" && url.pathname === "/api/admin/requests") return adminRequests(request, env);
      const decision = url.pathname.match(/^\/api\/admin\/requests\/([^/]+)\/(approve|reject)$/);
      if (request.method === "POST" && decision) return decideRequest(request, env, decodeURIComponent(decision[1]), decision[2] as "approve" | "reject");
      return errorResponse("Not found.", 404);
    } catch (error) {
      console.error("Licensing service request failed", error);
      return errorResponse("Unexpected service error.", 500);
    }
  },
};
