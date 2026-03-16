const { chromium } = require('playwright');
const path = require('path');
const http = require('http');
const fs = require('fs');
const { spawn } = require('child_process');

function loadConf() {
    const confPath = path.join(__dirname, '.conf');
    const lines = fs.readFileSync(confPath, 'utf8').split('\n');
    const conf = {};
    for (const line of lines) {
        const [key, ...rest] = line.split('=');
        if (key && rest.length) conf[key.trim()] = rest.join('=').trim();
    }
    return conf;
}

async function waitForPort(port, timeout = 10000) {
    const start = Date.now();
    while (Date.now() - start < timeout) {
        const ok = await new Promise(resolve => {
            http.get(`http://127.0.0.1:${port}/json/version`, res => {
                resolve(res.statusCode === 200);
            }).on('error', () => resolve(false));
        });
        if (ok) return true;
        await new Promise(r => setTimeout(r, 300));
    }
    return false;
}

function getTodayStr() {
    const now = new Date();
    return `${now.getFullYear()}/${now.getMonth() + 1}/${now.getDate()}`;
}

function getFileTimestamp() {
    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    return `${now.getFullYear()}${pad(now.getMonth()+1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
}

const KYUKA_KEYWORDS = ['全日有給休暇', '全日休暇', '午前半休', '午後半休', '午前休', '午後休', '半休', '全休', '欠勤'];
const ALLDAY_KYUKA   = ['全日有給休暇', '全日休暇', '全休', '欠勤'];
const HALF_KYUKA     = ['午前半休', '午後半休', '午前休', '午後休', '半休'];
const EXPLICIT_TIME_PATTERN = /(?:出社予定|時間|出勤|予定)[　\s:：]*(\d{1,2}:\d{2})/;
const DASH_TIME_PATTERN     = /[\-ー]{2}:[\-ー]{2}/;
const SKIP_SCHEDULE_KEYWORDS = [':休み', ':午前半休', ':午後半休'];

function findKyuka(text) {
    for (const k of KYUKA_KEYWORDS) {
        if (text.includes(k)) return k;
    }
    return null;
}

function findExplicitTime(text) {
    if (DASH_TIME_PATTERN.test(text)) return '--:--';
    const m = text.match(EXPLICIT_TIME_PATTERN);
    if (m) return m[1];
    const m2 = text.match(/⑤(\d{1,2}:\d{2})/);
    if (m2) return m2[1];
    return null;
}

// 全角数字を半角に変換
function toHalfWidth(str) {
    return str.replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
}

function parseEntry(entryText) {
    const lines = entryText.split('\n').map(l => toHalfWidth(l.trim())).filter(l => l);

    // ---- 形式①: 「N、値」または「N.値」形式 ----
    const hasReadtenFormat = lines.some(l => /^\d+[、,.]\s*\S/.test(l));
    if (hasReadtenFormat) {
        const map = {};
        for (const line of lines) {
            const m = line.match(/^(\d+)[、,.]\s*(.+)$/);
            if (m) map[parseInt(m[1])] = m[2].trim();
        }
        const name = map[2] || null;
        if (!name) return null;

        const kyuka = findKyuka(entryText) || null;
        const atsukai = map[3] || '';
        const riyu    = map[4] || '';

        if (!kyuka && !atsukai) return null;

        let time = '--:--';
        const timeVal = map[5] || '';
        if (/\d{1,2}:\d{2}/.test(timeVal)) {
            time = timeVal.match(/(\d{1,2}:\d{2})/)[1];
        } else if (kyuka && ALLDAY_KYUKA.includes(kyuka)) {
            time = '--:--';
        } else if (kyuka && HALF_KYUKA.includes(kyuka)) {
            time = '14:30';
        }

        return { name, time, kyuka: kyuka || '', atsukai, riyu };
    }

    // ---- 形式②: 丸数字・ラベル付き・漢字行 形式 ----
    const kyuka = findKyuka(entryText) || null;

    let atsukai = '';
    let riyu = '';

    const atsukaiCircle = entryText.match(/③([^①②③④⑤⑥\n]+)/);
    if (atsukaiCircle) atsukai = atsukaiCircle[1].trim();

    const riyuCircle = entryText.match(/④([^①②③④⑤⑥\n]+)/);
    if (riyuCircle) riyu = riyuCircle[1].trim();

    if (!atsukai) {
        const m = entryText.match(/^(?:連絡|扱い|届出処理)[\s　：:]+(.+)$/m);
        if (m) atsukai = m[1].trim();
    }
    if (!riyu) {
        const m = entryText.match(/^(?:事由|理由|事柄|簡易事由)[\s　：:]+(.+)$/m);
        if (m) riyu = m[1].trim();
    }

    if (!kyuka && !atsukai) return null;

    let time;
    const explicitTime = findExplicitTime(entryText);
    if (explicitTime) {
        time = explicitTime;
    } else if (kyuka && ALLDAY_KYUKA.includes(kyuka)) {
        time = '--:--';
    } else if (kyuka && HALF_KYUKA.includes(kyuka)) {
        time = '14:30';
    } else {
        time = '--:--';
    }

    let name = null;

    // 「②名前　○○」ラベル付き（丸数字あり）
    const nameLabelMatch = entryText.match(/[②2][.\s]?\s*名前[\s　：:]+([^\s　①②③④⑤⑥\n]+(?:\s+[^\s　①②③④⑤⑥\n]+)?)/);
    if (nameLabelMatch) name = nameLabelMatch[1].trim();

    // 「名前　○○」ラベルのみ（丸数字なし）
    if (!name) {
        const plainLabelMatch = entryText.match(/^名前[\s　：:]+(.+)$/m);
        if (plainLabelMatch) name = plainLabelMatch[1].trim();
    }

    // 「②○○」ラベルなし丸数字
    if (!name) {
        const circleTwoMatch = entryText.match(/②([^①②③④⑤⑥\n]{1,20})/);
        if (circleTwoMatch) {
            let candidate = circleTwoMatch[1].trim().replace(/^名前[\s　：:]+/, '').trim();
            if (candidate && !/^\d+(\.\d+)?$/.test(candidate) && !KYUKA_KEYWORDS.some(k => candidate.includes(k)) && !/\d{1,2}:\d{2}/.test(candidate)) {
                name = candidate;
            }
        }
    }

    // 漢字のみ行
    if (!name) {
        const originalLines = entryText.split('\n').map(l => l.trim()).filter(l => l);
        for (const line of originalLines) {
            if (/^\d+(\.\d+)?$/.test(line)) continue;
            if (/\d{4}\/\d+\/\d+/.test(line)) continue;
            if (KYUKA_KEYWORDS.some(k => line.includes(k))) continue;
            if (/\d{1,2}:\d{2}/.test(line)) continue;
            if (/[（(].+[）)]/.test(line)) continue;
            if (/^[①②③④⑤⑥]/.test(line)) continue;
            if (/^[\u3000-\u9FFF\u30A0-\u30FF\u3040-\u309F 　]{2,15}$/.test(line)) {
                name = line.trim();
                break;
            }
        }
    }

    if (!name) return null;
    return { name, time, kyuka: kyuka || '', atsukai, riyu };
}

// header文字列（例: "2026/3/16(月) 10:54"）から比較用の時刻数値を返す
function headerToMinutes(header) {
    const m = header.match(/(\d{1,2}):(\d{2})\s*$/);
    if (!m) return 0;
    return parseInt(m[1]) * 60 + parseInt(m[2]);
}

// 休暇種別から Event の select value を決定
function kyukaToEventValue(kyuka) {
    if (!kyuka) return null;
    if (ALLDAY_KYUKA.some(k => kyuka.includes(k))) return 's,:休み';
    if (kyuka.includes('午前')) return 's,:午前半休';
    if (kyuka.includes('午後')) return 's,:午後半休';
    if (kyuka.includes('半休') || kyuka.includes('半日')) return 's,:午前半休';
    return 's,:休み';
}

async function searchUserAndGetWeekView(page, setUrl, name) {
    await page.goto(setUrl, { waitUntil: 'domcontentloaded', timeout: 15000 }).catch(() => {});
    await page.waitForTimeout(1500);

    const nameParts = name.trim().split(/[\s　]+/);
    const candidates = [name];
    if (nameParts.length >= 2) {
        candidates.push(nameParts[0]);
        candidates.push(nameParts.slice(1).join(''));
    } else {
        const n = name.replace(/\s/g, '');
        if (n.length >= 4) {
            const half = Math.floor(n.length / 2);
            candidates.push(n.substring(0, half));
            candidates.push(n.substring(half));
        } else if (n.length === 3) {
            candidates.push(n.substring(0, 2));
            candidates.push(n.substring(1));
        }
    }

    for (const keyword of candidates) {
        if (!keyword) continue;

        const searchInputs = await page.locator('input[name="Text"]').all();
        const searchInput = searchInputs.length >= 2 ? searchInputs[1] : searchInputs[0];
        await searchInput.fill('');
        await searchInput.fill(keyword);
        await page.waitForTimeout(300);

        const searchBtn = page.locator('input[type="button"][value="ユーザー/施設検索"]').first();
        await searchBtn.click();
        await page.waitForTimeout(2000);

        const userInfo = await page.evaluate(() => {
            const rows = Array.from(document.querySelectorAll('tr'));
            const seen = new Set();
            const users = [];
            for (const row of rows) {
                const th = row.querySelector('th');
                if (!th) continue;
                const thText = (th.innerText || '').trim();
                if (!thText.includes('月予定')) continue;
                const userName = thText.split('\n')[0].trim();
                if (seen.has(userName)) continue;
                seen.add(userName);
                const cells = Array.from(row.querySelectorAll('td'));
                const todayCell = cells[0] ? (cells[0].innerText || '').trim() : '';
                users.push({ userName, todaySchedule: todayCell });
            }
            return users;
        });

        if (userInfo.length === 1) {
            return { keyword, userInfo: userInfo[0] };
        }
    }

    return null;
}

/**
 * グループ週画面の「+」ボタン（ScheduleEntry）をクリックして登録フォームを開き、
 * Event・Detail・時間を入力して登録する。
 * Detail には「時間/理由/扱い」形式で入力（休暇ありでも理由があれば含める）。
 */
async function registerSchedule(page, item, baseUrl) {
    // 「+」ボタンのリンク（ScheduleEntry）を取得
    const entryUrl = await page.evaluate((base) => {
        const links = Array.from(document.querySelectorAll('a[href*="ScheduleEntry"]'));
        for (const a of links) {
            const href = a.getAttribute('href');
            if (href) return href.startsWith('http') ? href : base + href.replace(/^.*ag\.cgi/, '');
        }
        return null;
    }, baseUrl);

    if (!entryUrl) {
        console.log('  ✗ ScheduleEntryリンクが見つかりません');
        return false;
    }

    await page.goto(entryUrl, { waitUntil: 'domcontentloaded', timeout: 15000 });
    await page.waitForTimeout(1500);

    // Detail欄に「時間/理由/扱い」を入力（休暇ありでも理由・扱いがあれば含める）
    const parts = [];
    if (item.time && item.time !== '--:--') parts.push(item.time);
    if (item.riyu) parts.push(item.riyu);
    if (item.atsukai) parts.push(item.atsukai);
    const detailText = parts.join('/');

    await page.locator('input[name="Detail"]').fill(detailText);

    // Event（予定種別）を設定
    const eventValue = kyukaToEventValue(item.kyuka);
    if (eventValue) {
        try {
            await page.selectOption('select[name="Event"]', { value: eventValue });
        } catch (e) {
            console.log(`  ⚠ Event選択スキップ: ${e.message}`);
        }
    }

    // 時間設定（--:-- 以外の場合のみ）
    if (item.time && item.time !== '--:--') {
        const [h, m] = item.time.split(':');
        try {
            await page.selectOption('select[name="SetTime.Hour"]', { value: String(parseInt(h)) });
            await page.selectOption('select[name="SetTime.Minute"]', { value: m });
        } catch (e) {
            console.log(`  ⚠ 時間設定スキップ: ${e.message}`);
        }
    }

    // 登録ボタンをクリック
    await page.locator('input[name="Entry"]').click();
    await page.waitForLoadState('domcontentloaded', { timeout: 10000 }).catch(() => {});
    await page.waitForTimeout(1000);

    return true;
}

async function run() {
    const conf = loadConf();
    const loginId = conf.ID       || '';
    const loginPw = conf.PASSWORD  || '';
    const loginUrl = conf.URL      || '';
    const pageUrl  = conf.PAGE_URL || '';
    const setUrl   = conf.SET_URL  || '';

    if (!loginId || !loginPw || !loginUrl || !pageUrl || !setUrl) {
        console.error('x .conf の設定を確認してください。');
        process.exit(1);
    }

    const u = new URL(loginUrl);
    const loginBase = `${u.protocol}//${u.host}/cgi-bin/cbag/ag.cgi`;
    const baseUrl   = `${u.protocol}//${u.host}/cgi-bin/cbag/`;

    console.log('ブラウザを起動中...');
    const chromiumPath = chromium.executablePath();
    const child = spawn(chromiumPath, [
        '--remote-debugging-port=19223',
        '--no-first-run',
        '--no-default-browser-check',
        loginBase
    ], { detached: true, stdio: 'ignore' });
    child.unref();

    const ready = await waitForPort(19223);
    if (!ready) { console.error('x タイムアウト'); process.exit(1); }

    const wsRes = await new Promise((resolve, reject) => {
        http.get('http://127.0.0.1:19223/json/version', res => {
            let body = '';
            res.on('data', c => body += c);
            res.on('end', () => resolve(JSON.parse(body).webSocketDebuggerUrl));
        }).on('error', reject);
    });

    const browser = await chromium.connectOverCDP(wsRes);
    const context = browser.contexts()[0];
    const page = context.pages()[0];
    await page.waitForTimeout(2000);

    // ---- ログイン ----
    const loginNameField  = page.locator('input[name="_account"]').first();
    const loginNameField2 = page.locator('input[type="text"]').first();
    let userField = (await loginNameField.count() > 0) ? loginNameField : loginNameField2;
    const passField = page.locator('input[type="password"]').first();

    if (userField && await passField.count() > 0) {
        await userField.fill(loginId);
        await passField.fill(loginPw);
        await page.waitForTimeout(300);
        const loginBtn = page.locator('input[type="submit"], button[type="submit"]').first();
        if (await loginBtn.count() > 0) await loginBtn.click();
        else await page.keyboard.press('Enter');
        await page.waitForLoadState('domcontentloaded', { timeout: 15000 }).catch(() => {});
        await page.waitForTimeout(2000);
    }

    // ---- 掲示板へ移動 ----
    console.log('掲示板ページへ移動中...');
    await page.goto(pageUrl, { waitUntil: 'domcontentloaded', timeout: 15000 }).catch(() => {});
    await page.waitForTimeout(2000);

    // ---- 当日の書き込みを取得 ----
    const todayStr = getTodayStr();
    console.log(`当日(${todayStr})の書き込みを検索中...`);

    const rawEntries = await page.evaluate((today) => {
        const results = [];
        const wrappers = document.querySelectorAll('div.vr_followWrapper');
        for (const wrapper of wrappers) {
            const timeEl = wrapper.querySelector('span.vr_followTime');
            if (!timeEl) continue;
            const timeText = (timeEl.innerText || '').trim();
            if (!timeText.includes(today)) continue;

            const ttEl = wrapper.querySelector('div.vr_followContents tt');
            if (!ttEl) continue;

            const bodyText = ttEl.innerHTML
                .replace(/<br\s*\/?>/gi, '\n')
                .replace(/<[^>]+>/g, '')
                .replace(/&nbsp;/g, ' ')
                .replace(/&amp;/g, '&')
                .replace(/&lt;/g, '<')
                .replace(/&gt;/g, '>')
                .trim();

            results.push({ date: timeText, body: bodyText });
        }
        return results;
    }, todayStr);

    if (rawEntries.length === 0) {
        console.log(`本日(${todayStr})の書き込みが見つかりませんでした。`);
        await browser.close();
        process.exit(0);
    }

    // ---- 重複除去・パース ----
    const seen = new Set();
    const parsedAll = [];

    for (const entry of rawEntries) {
        const key = entry.date + '|' + entry.body;
        if (seen.has(key)) continue;
        seen.add(key);

        const result = parseEntry(entry.body);
        if (!result) continue;
        if (/[（(].+[）)]/.test(result.name)) continue;
        parsedAll.push({ header: entry.date, ...result });
    }

    // ---- 同一人物の書き込みは最新のみ残す ----
    const latestMap = new Map();
    for (const item of parsedAll) {
        const existing = latestMap.get(item.name);
        if (!existing || headerToMinutes(item.header) > headerToMinutes(existing.header)) {
            latestMap.set(item.name, item);
        }
    }
    const parsed = Array.from(latestMap.values());

    if (parsed.length === 0) {
        console.log('抽出できる書き込みが見つかりませんでした。');
        await browser.close();
        process.exit(0);
    }

    // ---- ファイル保存 ----
    const listDir = path.join(__dirname, 'list');
    if (!fs.existsSync(listDir)) fs.mkdirSync(listDir);
    const timestamp = getFileTimestamp();
    const fileName  = `${timestamp}.txt`;
    const filePath  = path.join(listDir, fileName);

    let content = `取得日時: ${new Date().toLocaleString('ja-JP')}\n`;
    content += `対象日付: ${todayStr}\n`;
    content += '='.repeat(40) + '\n\n';
    for (const item of parsed) {
        content += `【${item.header}】\n`;
        content += `名前　: ${item.name}\n`;
        content += `時間　: ${item.time}\n`;
        content += `扱い　: ${item.atsukai}\n`;
        content += `理由　: ${item.riyu}\n`;
        content += `休暇　: ${item.kyuka}\n\n`;
    }
    fs.writeFileSync(filePath, content, 'utf8');

    console.log('\n--- 抽出結果 ---');
    for (const item of parsed) {
        const memo = [item.time, item.riyu, item.atsukai].filter(v => v && v !== '--:--').join('/');
        console.log(`名前: ${item.name}  時間: ${item.time}  休暇: ${item.kyuka || '（なし）'}  メモ: ${memo}`);
    }
    console.log(`保存先: list\\${fileName}`);

    // ---- ユーザー検索 & スケジュール確認 & 登録 ----
    console.log('\nユーザー検索 & スケジュール確認中...');
    const searchResults = [];

    for (const item of parsed) {
        const found = await searchUserAndGetWeekView(page, setUrl, item.name);
        if (!found) {
            searchResults.push({ ...item, skipReason: 'ユーザー特定不可' });
            continue;
        }

        const todaySchedule = found.userInfo.todaySchedule;
        const matchedKw = SKIP_SCHEDULE_KEYWORDS.find(kw => todaySchedule.includes(kw));

        // 既に休み/午前半休/午後半休 が登録済みでも、理由・扱いがあればDetailに入れて登録する
        // 理由も扱いもなければスキップ
        const hasDetail = item.riyu || item.atsukai;
        if (matchedKw && !hasDetail) {
            searchResults.push({ ...item, searchKeyword: found.keyword, skipReason: `登録済み・理由なし(${matchedKw})` });
        } else {
            console.log(`  → ${item.name} さんのスケジュールを登録中...`);
            const ok = await registerSchedule(page, item, baseUrl);
            if (ok) {
                const memo = [item.time, item.riyu, item.atsukai].filter(v => v && v !== '--:--').join('/');
                console.log(`  ✓ 登録完了 (Detail: ${memo})`);
                searchResults.push({ ...item, searchKeyword: found.keyword, skipReason: null, registered: true });
            } else {
                searchResults.push({ ...item, searchKeyword: found.keyword, skipReason: '登録失敗' });
            }
        }
    }

    // ---- 結果サマリー ----
    console.log('\n========================================');
    console.log('処理結果サマリー');
    console.log('========================================');

    const registered = searchResults.filter(r => !r.skipReason);
    const skipped    = searchResults.filter(r => r.skipReason);

    if (registered.length > 0) {
        console.log('\n【登録完了】');
        for (const r of registered) {
            const memo = [r.time, r.riyu, r.atsukai].filter(v => v && v !== '--:--').join('/');
            console.log(`  [登録済] ${r.name}  休暇: ${r.kyuka || '（なし）'}  メモ: ${memo}`);
        }
    } else {
        console.log('\n登録対象者はいません。');
    }

    if (skipped.length > 0) {
        console.log('\n【スキップ】');
        for (const r of skipped) {
            console.log(`  [スキップ] ${r.name} → ${r.skipReason}`);
        }
    }

    console.log(`\n合計: ${parsed.length}件 / 登録: ${registered.length}件 / スキップ: ${skipped.length}件`);

    await browser.close();

    console.log('\n========================================');
    console.log('[OK] スケジュール確認完了！');
    console.log('========================================');
}

run().catch(e => {
    console.error('Error:', e.message);
    process.exit(1);
});
