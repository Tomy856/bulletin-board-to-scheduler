const { chromium } = require('playwright');
const path = require('path');
const http = require('http');
const fs = require('fs');
const { spawn } = require('child_process');
const readline = require('readline');

// ---- 当日登録記録ファイルのパス ----
function getRegisteredFilePath(baseDir) {
    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    const dateKey = `${now.getFullYear()}${pad(now.getMonth()+1)}${pad(now.getDate())}`;
    return path.join(baseDir, 'list', `registered_${dateKey}.json`);
}

// 登録記録を読み込む（なければ空オブジェクト）
function loadRegistered(baseDir) {
    const filePath = getRegisteredFilePath(baseDir);
    if (!fs.existsSync(filePath)) return {};
    try {
        return JSON.parse(fs.readFileSync(filePath, 'utf8'));
    } catch { return {}; }
}

// 登録記録に1件追記保存
function saveRegistered(baseDir, name, kyuka, detail) {
    const filePath = getRegisteredFilePath(baseDir);
    const data = loadRegistered(baseDir);
    data[name] = { kyuka, detail, registeredAt: new Date().toLocaleString('ja-JP') };
    fs.writeFileSync(filePath, JSON.stringify(data, null, 2), 'utf8');
}


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

function askQuestion(query) {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout,
    });
    return new Promise(resolve => rl.question(query, ans => {
        rl.close();
        resolve(ans);
    }));
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
const EXPLICIT_TIME_PATTERN = /(?:出社予定|時間|出勤|予定)[　\s:：]*(\d{1,2}\s*[:\：]\s*\d{2})/;
const DASH_TIME_PATTERN     = /[\-ー]{2}:[\-ー]{2}/;
const SKIP_SCHEDULE_KEYWORDS = [':休み', ':午前半休', ':午後半休'];

function findKyuka(text) {
    for (const k of KYUKA_KEYWORDS) {
        if (text.includes(k)) return k;
    }
    return null;
}

// 「12時半」「12時30分」「12:30」「14 :00」などを "12:30" 形式に正規化する
function normalizeTimeStr(raw) {
    if (!raw) return null;
    // スペース入り「14 :00」「14: 00」なども HH:MM に正規化
    const colonMatch = raw.match(/(\d{1,2})\s*:\s*(\d{2})/);
    if (colonMatch) return `${colonMatch[1]}:${colonMatch[2].padStart(2,'0')}`;
    // 「12時半」→ 12:30
    const jiHanMatch = raw.match(/(\d{1,2})時半/);
    if (jiHanMatch) return `${jiHanMatch[1]}:30`;
    // 「12時30分」→ 12:30
    const jiMinMatch = raw.match(/(\d{1,2})時(\d{1,2})分/);
    if (jiMinMatch) return `${jiMinMatch[1]}:${String(jiMinMatch[2]).padStart(2,'0')}`;
    // 「12時」→ 12:00
    const jiMatch = raw.match(/(\d{1,2})時/);
    if (jiMatch) return `${jiMatch[1]}:00`;
    return null;
}

function findExplicitTime(text) {
    if (DASH_TIME_PATTERN.test(text)) return '--:--';
    const m = text.match(EXPLICIT_TIME_PATTERN);
    // ラベル付き時刻（スペース入り対応のため normalizeTimeStr に通す）
    if (m) return normalizeTimeStr(m[1]) || m[1];
    // ⑤ の後の時刻文字列を取得（「12:30」「12時半」「12時30分」に対応）
    const m2 = text.match(/⑤([^①②③④⑤⑥\n]{1,10})/);
    if (m2) {
        const normalized = normalizeTimeStr(m2[1].trim());
        if (normalized) return normalized;
    }
    // ラベルなし単独時刻 (14:00) を拾う
    const m3 = text.match(/(\d{1,2}\s*[:\：]\s*\d{2})/);
    if (m3) return normalizeTimeStr(m3[1]);
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

    let time = '--:--';
    const explicitTime = findExplicitTime(entryText);
    if (explicitTime) {
        time = explicitTime;
    } else if (kyuka && ALLDAY_KYUKA.includes(kyuka)) {
        time = '--:--';
    } else if (kyuka && HALF_KYUKA.includes(kyuka)) {
        time = '14:30';
    }

    let name = null;

    // 「②名前　○○」「②氏名　○○」ラベル付き（丸数字あり）
    const nameLabelMatch = entryText.match(/[②2][.\s]?\s*(?:名前|氏名)[\s　：:]+([^\s　①②③④⑤⑥\n]+(?:\s+[^\s　①②③④⑤⑥\n]+)?)/);
    if (nameLabelMatch) name = nameLabelMatch[1].trim();

    // 「名前　○○」「氏名　○○」ラベルのみ（丸数字なし）
    if (!name) {
        const plainLabelMatch = entryText.match(/^(?:名前|氏名)[\s　：:]+(.+)$/m);
        if (plainLabelMatch) name = plainLabelMatch[1].trim();
    }

    // 「②○○」ラベルなし丸数字
    if (!name) {
        const circleTwoMatch = entryText.match(/②([^①②③④⑤⑥\n]{1,20})/);
        if (circleTwoMatch) {
            let candidate = circleTwoMatch[1].trim().replace(/^(?:名前|氏名)[\s　：:]+/, '').trim();
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
                const todayCellEl = cells[0];
                const todayCellText = todayCellEl ? (todayCellEl.innerText || '').trim() : '';
                
                let registerUrl = null;
                let editUrl = null;
                let deleteUrl = null;
                let scheduleViewUrl = null;
                if (todayCellEl) {
                    const addLink = todayCellEl.querySelector('a[href*="ScheduleEntry"]');
                    if (addLink) registerUrl = addLink.getAttribute('href');
                    const editLink = todayCellEl.querySelector('a[href*="ScheduleEdit"]');
                    if (editLink) editUrl = editLink.getAttribute('href');
                    const deleteLink = todayCellEl.querySelector('a[href*="ScheduleDelete"]');
                    if (deleteLink) deleteUrl = deleteLink.getAttribute('href');

                    // 複数の予定がある中から、休暇関連のリンクを優先して探す
                    const detailLinks = Array.from(todayCellEl.querySelectorAll('a[href*="ScheduleView"], a[href*="ScheduleDetail"], a[href*="sid="]'));
                    const vacationKeywords = [':休み', ':午前半休', ':午後半休'];
                    let targetLink = detailLinks.find(a => {
                        const text = (a.innerText || '').trim();
                        return vacationKeywords.some(kw => text.includes(kw));
                    });
                    
                    // 休暇リンクが見つからなければ最初のリンクを代用
                    const viewLink = targetLink || detailLinks[0];
                    if (viewLink) scheduleViewUrl = viewLink.getAttribute('href');
                    
                    // デバッグ: すべてのリンクを記録
                    const allLinks = detailLinks.map(a => a.getAttribute('href')).join(' | ');
                    if (allLinks) console.log('TD links:', allLinks);
                }
                
                users.push({ userName, todaySchedule: todayCellText, registerUrl, editUrl, deleteUrl, scheduleViewUrl });
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
/**
 * 既存スケジュールを変更する（修正登録）
 * Attachment 1->2->3 のフローに従い、予定種別を更新し、時刻を解除（--:--）して保存する
 */
async function modifySchedule(page, item, baseUrl, isAllMode, scheduleViewUrl) {
    if (!scheduleViewUrl) return false;

    const viewUrl = scheduleViewUrl.startsWith('http') ? scheduleViewUrl : baseUrl + scheduleViewUrl.replace(/^.*ag\.cgi/, '');

    console.log(`  修正のため詳細画面へ移動: ${viewUrl}`);
    await page.goto(viewUrl, { waitUntil: 'domcontentloaded', timeout: 15000 }).catch(() => {});
    await page.waitForTimeout(1000);

    // 「変更する」リンクを探してクリック (Attachment 2 の操作)
    const modifyLink = page.locator('a[href*="ScheduleModify"], a:has-text("変更する")').first();
    if (await modifyLink.count() === 0) {
        console.log('  ✗ 変更ボタンが見つかりません');
        return false;
    }
    await modifyLink.click();
    await page.waitForLoadState('domcontentloaded', { timeout: 10000 }).catch(() => {});
    await page.waitForTimeout(1500);

    // 予定種別を設定 (Attachment 3 の操作)
    const eventValue = kyukaToEventValue(item.kyuka);
    if (eventValue) {
        try {
            await page.selectOption('select[name="Event"]', { value: eventValue });
        } catch (e) {
            console.log(`  ⚠ Event選択スキップ: ${e.message}`);
        }
    }

    // 時刻設定を解除する（すべて "--" に設定）
    // select の最初の選択肢（index:0）が通常「--」
    try {
        await page.selectOption('select[name="SetTime.Hour"]', { index: 0 }).catch(() => {});
        await page.selectOption('select[name="SetTime.Minute"]', { index: 0 }).catch(() => {});
        await page.selectOption('select[name="EndTime.Hour"]', { index: 0 }).catch(() => {});
        await page.selectOption('select[name="EndTime.Minute"]', { index: 0 }).catch(() => {});
    } catch (e) {
        console.log(`  ⚠ 時刻解除失敗: ${e.message}`);
    }

    // メモ（Detail）を更新
    const parts = [];
    if (item.time && item.time !== '--:--') parts.push(item.time);
    if (item.riyu) parts.push(item.riyu);
    if (item.atsukai) parts.push(item.atsukai);
    const detailText = parts.join('/');
    await page.locator('input[name="Detail"]').fill(detailText);

    // 登録確認（allオプションがない場合）
    if (!isAllMode) {
        console.log(`\n[修正内容の確認]`);
        console.log(`名前: ${item.name}`);
        console.log(`変更: ${item.kyuka || '休み'} (時刻解除: --:--)`);
        console.log(`内容: ${detailText}`);
        const answer = await askQuestion('この内容でスケジュールを変更しますか？ (y/n): ');
        if (answer.toLowerCase() !== 'y') {
            return 'cancelled';
        }
    }

    // 「変更する」ボタンをクリック
    const submitBtn = page.locator('input[type="submit"][value*="変更する"], input[name="Modify"], input[name="Entry"]').first();
    await submitBtn.click();
    await page.waitForLoadState('domcontentloaded', { timeout: 10000 }).catch(() => {});
    await page.waitForTimeout(1000);

    return true;
}

/**
 * 既存スケジュールを削除する（修正登録の前処理）
 * サイボウズの削除はScheduleDeleteページへGETするだけ
 */
async function deleteSchedule(page, baseUrl, deleteUrl, scheduleViewUrl) {
    // deleteUrl がなければ詳細画面から取得を試みる
    if (!deleteUrl && scheduleViewUrl) {
        console.log('  deleteUrl が null のため詳細画面から削除リンクを取得...');
        const viewUrl = scheduleViewUrl.startsWith('http') ? scheduleViewUrl : baseUrl + scheduleViewUrl.replace(/^.*ag\.cgi/, '');
        await page.goto(viewUrl, { waitUntil: 'domcontentloaded', timeout: 15000 }).catch(() => {});
        await page.waitForTimeout(1000);
        deleteUrl = await page.evaluate(() => {
            const link = document.querySelector('a[href*="ScheduleDelete"]');
            return link ? link.getAttribute('href') : null;
        });
        console.log(`  詳細画面から取得した deleteUrl: ${deleteUrl}`);
    }
    if (!deleteUrl) {
        console.log('  ✗ deleteUrl が取得できなかったため削除スキップ');
        return false;
    }
    const delUrl = deleteUrl.startsWith('http') ? deleteUrl : baseUrl + deleteUrl.replace(/^.*ag\.cgi/, '');
    console.log(`  削除URL: ${delUrl}`);
    await page.goto(delUrl, { waitUntil: 'domcontentloaded', timeout: 15000 }).catch(() => {});
    await page.waitForTimeout(1500);

    // 削除確認ページのボタン一覧をログ出力
    const btns = await page.evaluate(() =>
        Array.from(document.querySelectorAll('input[type=submit], input[type=button], button'))
            .map(b => `[${b.tagName}] name=${b.name} value=${b.value || b.innerText}`)
    );
    console.log(`  削除ページのボタン: ${btns.join(' / ')}`);

    // 削除確認ボタンを探してクリック（サイボウズの削除確認ページ対応）
    const confirmBtn = page.locator(
        'input[name="Delete"], input[value="削除する"], input[value="削除"], button:has-text("削除")'
    ).first();
    if (await confirmBtn.count() > 0) {
        await confirmBtn.click();
        await page.waitForLoadState('domcontentloaded', { timeout: 10000 }).catch(() => {});
        await page.waitForTimeout(1000);
        console.log('  削除ボタンをクリックしました');
    } else {
        console.log('  ⚠ 削除確認ボタンが見つかりませんでした（URLで直接削除された可能性あり）');
    }
    return true;
}

async function registerSchedule(page, item, baseUrl, isAllMode, registerUrl) {
    if (!registerUrl) {
        console.log('  ✗ 登録用(＋)リンクが見つかりません');
        return false;
    }

    const entryUrl = registerUrl.startsWith('http') ? registerUrl : baseUrl + registerUrl.replace(/^.*ag\.cgi/, '');

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

    // 登録確認（allオプションがない場合）
    if (!isAllMode) {
        console.log(`\n[登録内容の確認]`);
        console.log(`名前: ${item.name}`);
        console.log(`種類: ${item.kyuka || '（なし）'}`);
        console.log(`内容: ${detailText}`);
        const answer = await askQuestion('この内容でスケジュールを登録しますか？ (y/n): ');
        if (answer.toLowerCase() !== 'y') {
            return 'cancelled';
        }
    }

    // 登録ボタンをクリック
    await page.locator('input[name="Entry"]').click();
    await page.waitForLoadState('domcontentloaded', { timeout: 10000 }).catch(() => {});
    await page.waitForTimeout(1000);

    return true;
}

async function run() {
    const isAllMode = process.argv.includes('all');
    const __baseDir = __dirname;
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
    const baseUrl   = `${u.protocol}//${u.host}/cgi-bin/cbag/ag.cgi`;

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

    // ---- 当日の登録記録を読み込む ----
    const registeredToday = loadRegistered(__baseDir);

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
        const eventValue    = kyukaToEventValue(item.kyuka);
        const eventLabel    = eventValue ? eventValue.split(',')[1] : null; // 例: ":午前半休"
        const prevRecord    = registeredToday[item.name]; // 前回実行時の登録記録

        const isAlreadyRegistered = eventLabel && todaySchedule.includes(eventLabel);
        const prevEventLabel      = prevRecord ? kyukaToEventValue(prevRecord.kyuka)?.split(',')[1] : null;
        // 前回登録済み & 現在スケジュールに前回の予定種別がない → 手動削除されたとみなす
        const isManuallyDeleted   = prevRecord && prevEventLabel && !todaySchedule.includes(prevEventLabel);

        // 現在のスケジュールに別種別の休暇が入っているか（登録記録の有無に関係なく判定）
        // SKIP_SCHEDULE_KEYWORDS のうち、今回登録しようとしている種別以外のものが入っていれば修正対象
        const otherKyukaInSchedule = SKIP_SCHEDULE_KEYWORDS.filter(kw => kw !== eventLabel)
                                        .find(kw => todaySchedule.includes(kw));

        // 修正登録が必要なケース：
        //   ① 前回記録あり & 種別変化 & 前回種別がまだスケジュールにある
        //   ② 前回記録なし（または種別同じ）& 別種別の休暇がスケジュールに存在する
        const isKyukaChanged = (prevRecord && prevRecord.kyuka !== item.kyuka
                                    && prevEventLabel && todaySchedule.includes(prevEventLabel))
                               || (!isAlreadyRegistered && otherKyukaInSchedule);

        // 修正時に削除すべき既存の種別ラベル
        const deleteTargetLabel = prevEventLabel && todaySchedule.includes(prevEventLabel)
                                    ? prevEventLabel : otherKyukaInSchedule;

        const memo = [item.time, item.riyu, item.atsukai].filter(v => v && v !== '--:--').join('/');

        if (isManuallyDeleted) {
            // 11:00に登録 → 手動削除された → 再登録しない
            console.log(`  ⚠ ${item.name} さんは手動削除済みのためスキップ`);
            searchResults.push({ ...item, searchKeyword: found.keyword, skipReason: '手動削除済み' });

        } else if (isKyukaChanged) {
            // 休暇種別が変わった（例: 半休 → 休み）
            // 削除して再登録ではなく、詳細画面からの「変更」を優先する（休み > 午前半休 = 午後半休 の優先度で上書き）
            const fromLabel = prevRecord ? prevRecord.kyuka : (deleteTargetLabel || '既存の予定');
            console.log(`\n  → ${item.name} さんの予定を修正中... (${fromLabel} → ${item.kyuka})`);

            let result;
            if (found.userInfo.scheduleViewUrl) {
                // 詳細画面（Attachment 2）からの変更フロー（Attachment 3 へ遷移）
                result = await modifySchedule(page, item, baseUrl, isAllMode, found.userInfo.scheduleViewUrl);
            } else {
                // 変更リンクが特定できない場合は従来通り削除＆再登録
                await deleteSchedule(page, baseUrl, found.userInfo.deleteUrl, found.userInfo.scheduleViewUrl);
                const refound = await searchUserAndGetWeekView(page, setUrl, item.name);
                result = await registerSchedule(page, item, baseUrl, isAllMode, refound ? refound.userInfo.registerUrl : found.userInfo.registerUrl);
            }
            if (result === true) {
                saveRegistered(__baseDir, item.name, item.kyuka, memo);
                console.log(`  ✓ 修正登録完了 (Detail: ${memo})`);
                searchResults.push({ ...item, searchKeyword: found.keyword, skipReason: null, registered: true, modified: true });
            } else if (result === 'cancelled') {
                searchResults.push({ ...item, searchKeyword: found.keyword, skipReason: 'ユーザーによるキャンセル' });
            } else {
                searchResults.push({ ...item, searchKeyword: found.keyword, skipReason: '修正登録失敗' });
            }

        } else if (isAlreadyRegistered) {
            // 同じ予定種別がすでに登録済み → スキップ
            searchResults.push({ ...item, searchKeyword: found.keyword, skipReason: `登録済み(${eventLabel})` });

        } else {
            // 未登録 → 新規登録
            console.log(`\n  → ${item.name} さんのスケジュールを登録中...`);
            const result = await registerSchedule(page, item, baseUrl, isAllMode, found.userInfo.registerUrl);
            if (result === true) {
                saveRegistered(__baseDir, item.name, item.kyuka, memo);
                console.log(`  ✓ 登録完了 (Detail: ${memo})`);
                searchResults.push({ ...item, searchKeyword: found.keyword, skipReason: null, registered: true });
            } else if (result === 'cancelled') {
                console.log(`  - ${item.name} さんの登録をスキップしました (ユーザーによるキャンセル)`);
                searchResults.push({ ...item, searchKeyword: found.keyword, skipReason: 'ユーザーによるキャンセル' });
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
            const tag = r.modified ? '[修正]' : '[登録]';
            console.log(`  ${tag} ${r.name}  休暇: ${r.kyuka || '（なし）'}  メモ: ${memo}`);
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
