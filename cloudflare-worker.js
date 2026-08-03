const GITHUB_OWNER = "kaideyuj";
const GITHUB_REPO = "QCQuizData";
const GITHUB_BRANCH = "main";
const ALLOWED_ORIGIN = "https://kaideyuj.github.io";

function response(body, status = 200, origin = ALLOWED_ORIGIN) {
  return new Response(body === null ? null : JSON.stringify(body), {
    status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Access-Control-Allow-Origin": origin,
      "Access-Control-Allow-Headers": "Content-Type, X-Sync-Password",
      "Access-Control-Allow-Methods": "GET, PUT, OPTIONS",
      "Cache-Control": "no-store",
      "Vary": "Origin"
    }
  });
}

function timingSafeEqual(left, right) {
  const a = new TextEncoder().encode(left || "");
  const b = new TextEncoder().encode(right || "");
  let difference = a.length ^ b.length;
  const length = Math.max(a.length, b.length);

  for (let index = 0; index < length; index += 1) {
    difference |= (a[index % (a.length || 1)] || 0)
      ^ (b[index % (b.length || 1)] || 0);
  }

  return difference === 0;
}

async function profilePath(code) {
  const digest = await crypto.subtle.digest(
    "SHA-256",
    new TextEncoder().encode(code)
  );
  const hex = [...new Uint8Array(digest)]
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
  return `progress/${hex}.json`;
}

function githubHeaders(token) {
  return {
    "Accept": "application/vnd.github+json",
    "Authorization": `Bearer ${token}`,
    "X-GitHub-Api-Version": "2022-11-28",
    "User-Agent": "QCQuiz-Sync"
  };
}

async function readProfile(code, env) {
  const path = await profilePath(code);
  const url = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}`
    + `/contents/${path}?ref=${GITHUB_BRANCH}`;
  const result = await fetch(url, { headers: githubHeaders(env.GITHUB_TOKEN) });

  if (result.status === 404) {
    return null;
  }
  if (!result.ok) {
    throw new Error(`GitHub read failed (${result.status})`);
  }

  const file = await result.json();
  const text = new TextDecoder().decode(
    Uint8Array.from(atob(file.content.replace(/\s/g, "")), (char) => char.charCodeAt(0))
  );
  return { profile: JSON.parse(text), sha: file.sha, path };
}

async function writeProfile(code, profile, env) {
  const existing = await readProfile(code, env);
  const path = existing?.path || await profilePath(code);
  const url = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}`
    + `/contents/${path}`;
  const bytes = new TextEncoder().encode(JSON.stringify(profile));
  let binary = "";
  bytes.forEach((byte) => { binary += String.fromCharCode(byte); });

  const payload = {
    message: `Update progress for ${code}`,
    content: btoa(binary),
    branch: GITHUB_BRANCH
  };
  if (existing?.sha) {
    payload.sha = existing.sha;
  }

  const result = await fetch(url, {
    method: "PUT",
    headers: {
      ...githubHeaders(env.GITHUB_TOKEN),
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });

  if (!result.ok) {
    throw new Error(`GitHub write failed (${result.status})`);
  }
}

export default {
  async fetch(request, env) {
    const origin = request.headers.get("Origin") || ALLOWED_ORIGIN;
    if (origin !== ALLOWED_ORIGIN) {
      return response({ error: "Origin not allowed" }, 403, ALLOWED_ORIGIN);
    }
    if (request.method === "OPTIONS") {
      return response(null, 204, origin);
    }

    const url = new URL(request.url);
    const match = url.pathname.match(/^\/progress\/([^/]+)$/);
    if (!match || !["GET", "PUT"].includes(request.method)) {
      return response({ service: "QCQuiz Sync", status: "ok" }, 200, origin);
    }
    if (!timingSafeEqual(request.headers.get("X-Sync-Password"), env.SYNC_PASSWORD)) {
      return response({ error: "同步密碼錯誤" }, 401, origin);
    }

    const code = decodeURIComponent(match[1]).trim();
    if (!code.startsWith("yuj-") || code.length > 30) {
      return response({ error: "無效的同步代號" }, 400, origin);
    }

    try {
      if (request.method === "GET") {
        const stored = await readProfile(code, env);
        return stored
          ? response({ profile: stored.profile }, 200, origin)
          : response({ error: "尚無雲端紀錄" }, 404, origin);
      }

      const body = await request.json();
      if (!body?.profile || body.profile.code !== code) {
        return response({ error: "進度資料無效" }, 400, origin);
      }
      await writeProfile(code, body.profile, env);
      return response({ saved: true }, 200, origin);
    } catch (error) {
      return response({ error: "雲端同步失敗", detail: error.message }, 502, origin);
    }
  }
};
