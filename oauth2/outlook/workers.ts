export default {
  async fetch(request: Request, env: any, ctx: ExecutionContext): Promise<Response> {
    const url = new URL(request.url);
    const path = url.pathname;
    let body;
    try {
      body = await request.json();
    } catch {
      return new Response("", { status: 204 }); // Không phản hồi nếu JSON không hợp lệ
    }

    const apiKey = body.api_key;
    if (!apiKey) return new Response("", { status: 204 });

    const keyCheck = await env.KV_CACHE.get(`outlook/api_key:${apiKey}`);
    if (!keyCheck) return new Response("", { status: 204 });

    if (path.endsWith("/refresh")) return xuLy_Token_Outlook(body, env, ctx);
    if (path.endsWith("/messenger")) return xuLy_Messenger_Outlook(body, env, ctx);

    return new Response("404 Not Found", { status: 404 });
  },

  scheduled // <-- Gán lại hàm đã export ở trên
};




// ✅ Cron xử lý API key hết hạn (Trigger gọi hàm này mỗi 30 phút)
export async function scheduled(controller, env, ctx) {
  const now = getTimeVietnam();
  const currentTime = convertVietnamTimeToDate(now);
  const keys = await env.KV_CACHE.list({ prefix: "outlook/api_key:" });

  await Promise.all(keys.keys.map(async (key) => {
    const expiryTimeStr = await env.KV_CACHE.get(key.name);
    if (!expiryTimeStr) return;
    let expiryTime;
    try { expiryTime = convertVietnamTimeToDate(expiryTimeStr); }
    catch { return; }
    if (expiryTime <= currentTime) {
      const expiredKey = key.name.replace("outlook/api_key:", "outlook/expired/api_key:");
      const value = await env.KV_CACHE.get(key.name);
      await Promise.all([
        env.KV_CACHE.put(expiredKey, value),
        env.KV_CACHE.delete(key.name)
      ]);
    }
  }));
}



// Chuẩn hóa thời gian Việt Nam
function getTimeVietnam() {
  const now = new Date();
  const vnOffset = 7 * 60; // phút
  const vnTime = new Date(now.getTime() + vnOffset * 60000);
  const pad = (n) => n.toString().padStart(2, "0");
  return `${pad(vnTime.getDate())}/${pad(vnTime.getMonth() + 1)}/${vnTime.getFullYear()} ` +
         `${pad(vnTime.getHours())}:${pad(vnTime.getMinutes())}:${pad(vnTime.getSeconds())}`;
}

function convertVietnamTimeToDate(vnTime: string): Date {
  const [dd, MM, yyyyHHmmss] = vnTime.split("/");
  const [yyyy, HHmmss] = yyyyHHmmss.split(" ");
  const [HH, mm, ss] = HHmmss.split(":");
  return new Date(`${yyyy}-${MM}-${dd}T${HH}:${mm}:${ss}`);
}



// ✅ Hàm 1: Lưu email vào KV_CACHE
export async function luu_Email_KV_OUTLOOK(apiKey: string, email: string, dataMoi: any, env: any): Promise<void> {
  // Không ghi nếu thiếu email hoặc pass!
  if (!email || !dataMoi?.pass) return;

  const keyData = `user:${apiKey}:${email.toLowerCase()}`;
  const keyIndex = `index:${email.toLowerCase()}`;
  const emailLC = email.toLowerCase();

  const dataCuStr = await env.KV_OUTLOOK.get(keyData);
  const thoiGian = getTimeVietnam();

  // Nếu chưa có key → ghi mới toàn bộ (yêu cầu refresh_token có giá trị)
  if (!dataCuStr) {
    if (!dataMoi.refresh_token) return; // Không ghi nếu không có refresh_token
    await Promise.all([
      env.KV_OUTLOOK.put(keyData, JSON.stringify({
        pass: dataMoi.pass || "",
        refresh_token: dataMoi.refresh_token,
        access_token: dataMoi.access_token || "",
        client_id: dataMoi.client_id || "",
        status_token: "Live",
        time_token: thoiGian
      })),
      env.KV_INDEX_OUTLOOK.put(keyIndex, `user:${apiKey}`)
    ]);
    return;
  }

  // Nếu đã có key → so sánh và cập nhật từng trường nếu thay đổi
  const dataCu = JSON.parse(dataCuStr);
  let coThayDoi = false;
  const dataCapNhat = { ...dataCu };

  for (const truong of ["pass", "client_id", "access_token"]) {
    if (dataMoi[truong] && dataMoi[truong] !== dataCu[truong]) {
      dataCapNhat[truong] = dataMoi[truong];
      coThayDoi = true;
    }
  }

  if (dataMoi.refresh_token && dataMoi.refresh_token !== dataCu.refresh_token) {
    dataCapNhat.refresh_token = dataMoi.refresh_token;
    dataCapNhat.time_token = thoiGian;
    dataCapNhat.status_token = "Live";
    coThayDoi = true;
  }

  if (!dataMoi.refresh_token) {
    dataCapNhat.status_token = "lock";
    dataCapNhat.time_token = thoiGian;
    coThayDoi = true;
  }

  if (!coThayDoi) return;

  await Promise.all([
    env.KV_OUTLOOK.put(keyData, JSON.stringify(dataCapNhat)),
    env.KV_INDEX_OUTLOOK.put(keyIndex, `user:${apiKey}`)
  ]);
}





// ✅ Xử lý Token Outlook
export async function xuLy_Token_Outlook(thamSo: any, env: any, ctx: ExecutionContext): Promise<Response> {
  const {
    api_key = "",
    server = "false",
    email = "",
    pass = "",
    refresh_token = "",
    client_id = ""
  } = thamSo || {};

  const dungServer = server === "true";
  const emailLC = email.toLowerCase();

  let tokenMoi = refresh_token;
  let idClient = client_id;

  if (dungServer && (!tokenMoi || !idClient)) {
    const duLieu = await env.KV_CACHE.get(`outlook/${api_key}:${emailLC}`);
    if (duLieu) {
      const parsed = JSON.parse(duLieu);
      if (parsed.pass !== pass) return new Response("", { status: 204 });
      tokenMoi = parsed.refresh_token || tokenMoi;
      idClient = parsed.client_id || idClient;
    }
  }

  const tokenRes = await lam_Moi_Token_Outlook({ refresh_token: tokenMoi, client_id: idClient }, env);
  const tokenData = await tokenRes.json();
  const token = tokenData.access_token || "";
  tokenMoi = tokenData.refresh_token || tokenMoi;
  const thoiGian = getTimeVietnam();

  if (!token) {
    const thongBao = dich_Loi_Microsoft(tokenData);
    const ketQuaPhanHoi = {
      status: thongBao.status,
      messenger: thongBao.messenger,
      email,
      pass,
      time_token: thoiGian
    };

    if (dungServer) {
      ctx.waitUntil(luu_Email_KV_OUTLOOK(api_key, emailLC, {
        pass,
        refresh_token: tokenMoi,
        access_token: "",
        client_id: idClient,
        time_token: thoiGian,
        status_token: thongBao.status === "lock" ? "lock" : "Live"
      }, env));
    }

    return new Response(JSON.stringify(ketQuaPhanHoi), { status: 200 });
  }

  const ketQuaPhanHoi = {
    status: "success",
    messenger: "Lấy access_token thành công",
    email,
    pass,
    refresh_token: tokenMoi,
    access_token: token,
    client_id: idClient,
    time_token: thoiGian
  };

  if (dungServer) {
    ctx.waitUntil(luu_Email_KV_OUTLOOK(api_key, emailLC, {
      pass,
      refresh_token: tokenMoi,
      access_token: token,
      client_id: idClient,
      time_token: thoiGian,
      status_token: "Live"
    }, env));
  }

  return new Response(JSON.stringify(ketQuaPhanHoi), { status: 200 });
}



// ✅ Làm mới token Outlook
export async function lam_Moi_Token_Outlook(thamSo: any, env: any): Promise<Response> {
  const {
    refresh_token = "",
    client_id = ""
  } = thamSo || {};

  const v2Res = await fetch("https://login.microsoftonline.com/common/oauth2/v2.0/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id,
      grant_type: "refresh_token",
      refresh_token,
      scope: "https://graph.microsoft.com/.default"
    })
  });

  const v2Data = await v2Res.json();
  if (v2Data.access_token) return new Response(JSON.stringify({ ...v2Data, version_used: "v2" }), { status: 200 });

  const v1Res = await fetch("https://login.microsoftonline.com/common/oauth2/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id,
      grant_type: "refresh_token",
      refresh_token,
      resource: "https://graph.microsoft.com"
    })
  });

  const v1Data = await v1Res.json();
  return new Response(JSON.stringify({ ...v1Data, version_used: "v1" }), { status: 200 });
}




// ✅ Xử lý Messenger Outlook
export async function xuLy_Messenger_Outlook(thamSo: any, env: any, ctx: ExecutionContext): Promise<Response> {
  const {
    api_key = "",
    server = "false",
    email = "",
    pass = "",
    refresh_token = "",
    client_id = "",
    access_token = "",
    reading = "",
    keywords = ""
  } = thamSo || {};

  const dungServer = server === "true";
  const emailLC = email.toLowerCase();
  const KV_KEY = `outlook/${api_key}:${emailLC}`;
  let token = access_token || "";
  let tokenMoi = refresh_token || "";
  let idClient = client_id || "";
  let time_token = "";

  // 1. ƯU TIÊN TOKEN TRONG REQUEST
  if (!token) {
    if (tokenMoi && idClient) {
      // Làm mới access_token từ refresh_token+client_id trong request
      const tokenRes = await lam_Moi_Token_Outlook({ refresh_token: tokenMoi, client_id: idClient }, env);
      const tokenData = await tokenRes.json();
      token = tokenData.access_token || "";
      tokenMoi = tokenData.refresh_token || tokenMoi;
      time_token = getTimeVietnam();
      if (!token) {
        const thongBao = dich_Loi_Microsoft(tokenData);
        return new Response(JSON.stringify({
          status: thongBao.status,
          messenger: thongBao.messenger,
          email,
          pass,
          refresh_token: tokenMoi,
          access_token: "",
          client_id: idClient,
          time_token
        }), { status: 200 });
      }
    }
    // 2. Fallback lấy từ KV nếu server:true (khi request không có đủ)
    else if (dungServer) {
      const duLieu = await env.KV_CACHE.get(KV_KEY);
      if (!duLieu) {
        return new Response(JSON.stringify({
          status: "fail",
          messenger: "Không tìm thấy dữ liệu trong cache",
          email,
          pass,
          refresh_token: "",
          access_token: "",
          client_id: "",
          time_token: getTimeVietnam()
        }), { status: 200 });
      }
      const parsed = JSON.parse(duLieu);
      if (parsed.pass !== pass) return new Response("", { status: 204 });

      token = parsed.access_token || "";
      tokenMoi = parsed.refresh_token || "";
      idClient = parsed.client_id || "";
      time_token = parsed.time_token || "";

      // Kiểm tra TTL access_token: còn hiệu lực < 1H thì dùng luôn, quá hạn thì làm mới
      let validToken = false;
      if (token && time_token) {
        try {
          const tokenTime = convertVietnamTimeToDate(time_token).getTime();
          if (Date.now() - tokenTime < 60 * 60 * 1000) validToken = true;
        } catch {}
      }
      if (!validToken && tokenMoi && idClient) {
        // Token hết hạn, làm mới rồi ghi lại cache sau
        const tokenRes = await lam_Moi_Token_Outlook({ refresh_token: tokenMoi, client_id: idClient }, env);
        const tokenData = await tokenRes.json();
        token = tokenData.access_token || "";
        tokenMoi = tokenData.refresh_token || tokenMoi;
        time_token = getTimeVietnam();
        if (!token) {
          const thongBao = dich_Loi_Microsoft(tokenData);
          return new Response(JSON.stringify({
            status: thongBao.status,
            messenger: thongBao.messenger,
            email,
            pass,
            refresh_token: tokenMoi,
            access_token: "",
            client_id: idClient,
            time_token
          }), { status: 200 });
        }
      }
      // Nếu vẫn không có access_token hợp lệ
      if (!token) {
        return new Response(JSON.stringify({
          status: "fail",
          messenger: "Không có access_token hoặc refresh_token hợp lệ",
          email,
          pass,
          refresh_token: tokenMoi,
          access_token: "",
          client_id: idClient,
          time_token
        }), { status: 200 });
      }
    }
  }

  // 3. Đọc mail bằng access_token này
  let code = null;
  let mailId = "";
  let mailBodyText = "";

  // Thử đọc mail 1 lần, nếu lỗi xác thực access_token thì thử làm mới và đọc lại (chỉ 1 lần)
  let tokenRefreshed = false;
  let thoiGian = getTimeVietnam();
  let docRes = await doc_Mail_Outlook({ access_token: token, filterUnread: reading === "tick" }, env);
  let data = await docRes.json();

  if (isAccessTokenError(data) && tokenMoi && idClient && !tokenRefreshed) {
    // Chỉ refresh 1 lần nếu lỗi xác thực
    tokenRefreshed = true;
    const tokenRes = await lam_Moi_Token_Outlook({ refresh_token: tokenMoi, client_id: idClient }, env);
    const tokenData = await tokenRes.json();
    token = tokenData.access_token || "";
    tokenMoi = tokenData.refresh_token || tokenMoi;
    thoiGian = getTimeVietnam();
    if (token) {
      docRes = await doc_Mail_Outlook({ access_token: token, filterUnread: reading === "tick" }, env);
      data = await docRes.json();
    } else {
      // Làm mới vẫn fail
      const thongBao = dich_Loi_Microsoft(tokenData);
      return new Response(JSON.stringify({
        status: thongBao.status,
        messenger: thongBao.messenger,
        email,
        pass,
        refresh_token: tokenMoi,
        access_token: "",
        client_id: idClient,
        time_token: thoiGian
      }), { status: 200 });
    }
  }

  // Nếu vẫn lỗi sau khi thử refresh, trả lỗi rõ ràng
  if (data?.error || !docRes.ok) {
    const thongBao = dich_Loi_Microsoft(data);
    return new Response(JSON.stringify({
      status: thongBao.status,
      messenger: thongBao.messenger,
      email,
      pass,
      refresh_token: tokenMoi,
      access_token: "",
      client_id: idClient,
      time_token: thoiGian
    }), { status: 200 });
  }

  // Duyệt mail lấy code
  const mails = data?.value || [];
  for (const item of mails) {
    const fromAddress = item.from?.emailAddress?.address?.toLowerCase() || "";
    if (!fromAddress.includes(keywords.toLowerCase())) continue;

    const body = item.body?.content || "";
    const plainBody = body.replace(/<[^>]+>/g, "").replace(/&[a-z]+;/gi, " ").trim();
    const match = plainBody.match(/\b\d{4,8}\b/);

    if (match) {
      code = match[0];
      mailId = item.id;
      mailBodyText = plainBody;
      break;
    }
  }

  // Chuẩn bị phản hồi
  thoiGian = getTimeVietnam();
  const ketQuaPhanHoi: any = {
    status: code ? "success" : "fail",
    messenger: code ? "Lấy mã xác minh thành công" : (mails.length ? "Không tìm thấy mã xác minh" : "Không tìm thấy email"),
    email,
    pass,
    refresh_token: tokenMoi,
    access_token: token,
    client_id: idClient,
    time_token: thoiGian
  };
  if (code) ketQuaPhanHoi.body = mailBodyText;
  ketQuaPhanHoi.code = code;

  // 4. Cập nhật lại cache nếu có server: true (ghi lại record mới nhất)
ctx.waitUntil((async () => {
  if (dungServer) {
    await luu_Email_KV_OUTLOOK(api_key, emailLC, {
      pass,
      refresh_token: tokenMoi,
      access_token: token,
      client_id: idClient,
      time_token: thoiGian,
      status_token: "Live"
    }, env);
  }
  if (mailId && reading === "tick") {
    // Delay 1-3 giây ngẫu nhiên (ví dụ)
    await new Promise(r => setTimeout(r, 1000 + Math.floor(Math.random() * 2000)));
    await fetch(`https://graph.microsoft.com/v1.0/me/messages/${mailId}`, {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ isRead: true })
    });
  }
  if (mailId && reading === "delete") {
    // Delay 1-3 giây ngẫu nhiên (nếu muốn)
    await new Promise(r => setTimeout(r, 1000 + Math.floor(Math.random() * 2000)));
    await fetch(`https://graph.microsoft.com/v1.0/me/messages/${mailId}`, {
      method: "DELETE",
      headers: { Authorization: `Bearer ${token}` }
    });
  }
})());

  return new Response(JSON.stringify(ketQuaPhanHoi), { status: 200 });
}

// === Hàm phụ: Xác định lỗi access_token hết hạn hoặc invalid ===
function isAccessTokenError(data: any): boolean {
  // Dựa vào code lỗi chuẩn của Microsoft
  const msg = (data?.error?.message || data?.error_description || data?.message || "").toLowerCase();
  return (
    msg.includes("access token") ||
    msg.includes("expired") ||
    msg.includes("invalid") ||
    msg.includes("authentication") ||
    msg.includes("unauthorized") ||
    msg.includes("token has expired") ||
    msg.includes("bearer") ||
    (data?.error?.code && (data.error.code === "InvalidAuthenticationToken" || data.error.code === "InvalidToken"))
  );
}






// ✅ Đọc mail Outlook từ inbox, chưa đọc, trong 15 phút gần nhất
export async function doc_Mail_Outlook(thamSo: any, env: any): Promise<Response> {
  const { access_token = "", filterUnread = false } = thamSo || {};
  if (!access_token) {
    return new Response(JSON.stringify({ value: [] }), { status: 200 });
  }

  const fifteenMinutesAgo = new Date(Date.now() - 15 * 60 * 1000).toISOString();
  let filter = `receivedDateTime ge ${fifteenMinutesAgo}`;
  if (filterUnread) filter = `isRead eq false and ${filter}`;

  const url = `https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages` +
    `?$top=5&$orderby=receivedDateTime desc` +
    `&$select=id,from,subject,receivedDateTime,body,isRead` +
    `&$filter=${encodeURIComponent(filter)}`;

  try {
    const res = await fetch(url, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${access_token}`,
        "Prefer": "outlook.body-content-type=\"text\""
      }
    });
    const data = await res.json();
    return new Response(JSON.stringify(data), { status: 200 });
  } catch (err) {
    return new Response(JSON.stringify({ value: [] }), { status: 200 });
  }
}












// ✅ Trích mã xác minh từ nội dung mail
export function extract_Code_Mail(noiDung: string, key: string): string | null {
  const plainText = noiDung.replace(/<[^>]+>/g, "").replace(/&[a-z]+;/gi, " ").trim();
  const keyLower = key.toLowerCase();
  if (["facebook", "facebookmail.com"].includes(keyLower)) {
    const match = plainText.match(/FB-\d{4,8}/i);
    return match ? match[0] : null;
  }
  const match = plainText.match(/\b\d{4,8}\b/);
  return match ? match[0] : null;
}



// Dịch mã lỗi Microsoft chuẩn hóa
export function dich_Loi_Microsoft(data: any): { status: "success" | "lock" | "fail", messenger: string } {
  const raw = (data?.error_description || data?.error?.message || data?.message || "").toLowerCase();

  if (!raw) return { status: "fail", messenger: "Không rõ lỗi từ máy chủ Microsoft." };

  const isLock = [
    "aadsts50076", "aadsts50079", "aadsts65001", "authorization denied",
    "aadsts70007", "aadsts54005", "invalid_grant", "locked", "disabled", "account is locked", "user account is locked"
  ].some(code => raw.includes(code));

  if (isLock) return { status: "lock", messenger: traThongBao(raw) };
  return { status: "fail", messenger: traThongBao(raw) };
}


function traThongBao(raw: string): string {
  if (raw.includes("aadsts70000")) return "Token không hợp lệ hoặc đã hết hạn.";
  if (raw.includes("aadsts70002")) return "Sai client_id hoặc client_secret.";
  if (raw.includes("aadsts70007")) return "Refresh token đã hết hạn.";
  if (raw.includes("aadsts70008")) return "Sai kiểu grant_type.";
  if (raw.includes("aadsts70011")) return "Scope không hợp lệ hoặc thiếu.";
  if (raw.includes("aadsts7000215")) return "Client secret không hợp lệ.";
  if (raw.includes("aadsts700016")) return "Ứng dụng không tồn tại (client_id sai).";
  if (raw.includes("aadsts7000218")) return "Định danh tài khoản không hợp lệ.";
  if (raw.includes("aadsts70020")) return "Request body không hợp lệ.";
  if (raw.includes("aadsts9002313")) return "Request thiếu hoặc sai tham số.";
  if (raw.includes("aadsts50076")) return "Yêu cầu xác thực hai bước (MFA).";
  if (raw.includes("aadsts50079")) return "Tài khoản chưa hoàn tất xác minh bảo mật.";
  if (raw.includes("aadsts50126")) return "Sai tài khoản hoặc mật khẩu Outlook.";
  if (raw.includes("aadsts50133")) return "Access token không hợp lệ.";
  if (raw.includes("aadsts65001")) return "Tài khoản chưa cấp quyền cho ứng dụng.";
  if (raw.includes("aadsts65004")) return "Cần quyền quản trị viên để truy cập.";
  if (raw.includes("aadsts53003")) return "Bị chặn bởi chính sách đăng nhập của tổ chức.";
  if (raw.includes("aadsts50020")) return "Tài khoản không thuộc tenant ứng dụng.";
  if (raw.includes("aadsts50059")) return "Tenant ID không hợp lệ hoặc không tồn tại.";
  if (raw.includes("aadsts54005")) return "Token đã hết hạn do người dùng không hoạt động.";
  if (raw.includes("aadsts500113")) return "Redirect URI không khớp.";
  if (raw.includes("invalid_grant")) return "Refresh token đã bị thu hồi hoặc không hợp lệ.";
  if (raw.includes("invalid_client")) return "Client ID hoặc secret không hợp lệ.";
  if (raw.includes("authorization denied")) return "Người dùng đã từ chối cấp quyền.";
  if (raw.includes("unauthorized")) return "Token hết hạn hoặc bị từ chối.";
  if (raw.includes("invalid_request")) return "Request không hợp lệ hoặc thiếu tham số.";
  if (raw.includes("bad request")) return "Yêu cầu không hợp lệ.";
  return "❌ Lỗi không xác định từ Microsoft.";
}
