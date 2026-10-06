/* ============================================================
 * ระบบแจ้งลา + โอที — functions/tg.js — ตัวส่งต่อ Telegram webhook (Cloudflare Pages Function) : https://<เว็บของคุณ>/tg
 *
 * ทำไมต้องมี: Telegram ยิงเข้า Google Apps Script ตรง ๆ ได้ แต่ Apps Script ตอบช้าและตอบแบบ redirect ทำให้ Telegram คิดว่าส่งไม่สำเร็จ
 * แล้วยิงซ้ำ/คิวค้าง (ปุ่มในบอทจะช้าหรือเงียบ) ตัวนี้ตอบ Telegram ทันที แล้วส่งต่อให้ Apps Script
 * พร้อมแนบรหัสลับที่ Telegram ส่งมาในหัวคำขอ ให้ Apps Script ตรวจ (ค่า tg_webhook_secret ในชีต Settings)
 *
 * เปิดใช้: 1) ใส่ URL /exec ของ Apps Script ที่ GAS_URL ด้านล่าง (หรือตั้งตัวแปร GAS_URL ในโปรเจกต์ Cloudflare Pages)
 *          2) ใส่ https://<เว็บของคุณ>/tg ในชีต Settings แถว tg_webhook_url  3) รันฟังก์ชัน connectTelegramNow ใน Apps Script
 * ============================================================ */
const GAS_URL = '';   // ← URL เดียวกับ API_URL ใน config.js

export async function onRequestPost(context) {
  const { request, env } = context;
  const secret = request.headers.get('X-Telegram-Bot-Api-Secret-Token') || '';
  if (!secret) return new Response('forbidden', { status: 403 });   // ไม่ได้มาจาก Telegram (ไม่มีรหัสลับ)
  const body = await request.text();
  if (body.length > 200000) return new Response('too large', { status: 413 });
  const target = (env && env.GAS_URL) || GAS_URL;
  if (!target) return new Response('GAS_URL not set', { status: 500 });
  context.waitUntil(
    fetch(target + '?tgs=' + encodeURIComponent(secret), {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body, redirect: 'follow'
    }).catch(() => {})
  );
  return new Response('ok');
}

export function onRequestGet() {
  return new Response('telegram relay: ok', { headers: { 'Content-Type': 'text/plain; charset=utf-8' } });
}
