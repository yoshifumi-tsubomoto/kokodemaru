/**
 * 採点アプリ 生徒ダッシュボード（Google Apps Script のWebアプリ）
 *
 * 公開の設定：「次のユーザーとして実行：自分（先生）」「アクセスできるユーザー：学校のドメインの全員」
 * - 先生（このスクリプトの持ち主）が開くと、先生用の画面（結果のZIPの取り込み・確認）
 * - 生徒が開くと、ログインしているGoogleアカウントの本人の結果だけを表示
 * 答案の画像と結果は、先生のGoogleドライブの専用フォルダに保存する（生徒とは共有しない）。
 *
 * ココデマル ダッシュボード（下の APP_VERSION の版）
 * Copyright 2026 坪本 吉史
 * Apache License 2.0 で配っています（https://www.apache.org/licenses/LICENSE-2.0）。
 * 自由にコピー・改変・再配布できます（有料でも可）。改変したものを配るときは、
 * この版の番号とこの文を残し、どこを変えたかを書いてください。
 */

// 版の番号（採点アプリの app/package.json の version と同じにする。ずれるとテストで分かる）
const APP_VERSION = '1.0.0'
const ROOT_FOLDER_NAME = '採点アプリ_ダッシュボード（生徒と共有しないでください）'
const MANIFEST = 'manifest.json'

// ---------- 画面 ----------

function doGet(e) {
  const params = (e && e.parameter) || {}
  const teacher = isTeacher_()
  // 先生は ?view=student&as=メールアドレス で、その生徒の画面を確認できる
  const page = teacher && params.view !== 'student' ? 'admin' : 'student'
  const t = HtmlService.createTemplateFromFile(page)
  // 画面のプログラムに埋め込む値は、JavaScriptの文字列として正しく書き出す
  t.previewAsJs = jsString_(teacher && page === 'student' ? String(params.as || '') : '')
  t.appUrlJs = jsString_(ScriptApp.getService().getUrl() || '')
  return t.evaluate()
    .setTitle(page === 'admin' ? 'テストの結果（先生用）' : 'テストの結果')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
}

/** HTML の <script> の中に安全に埋め込める文字列（"…" の形。</script> などで途切れないようにする） */
function jsString_(s) {
  return JSON.stringify(String(s)).replace(/</g, '\\u003c')
}

/** HTML の中から共通の部品を読み込む */
function include_(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent()
}

// ---------- 共通 ----------

function me_() {
  return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase()
}

function isTeacher_() {
  const owner = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase()
  return !!owner && me_() === owner
}

function assertTeacher_() {
  if (!isTeacher_()) throw new Error('先生だけが使える操作です')
}

/** 保存先のフォルダ（なければ作る。フォルダの番号はスクリプトのプロパティに覚える） */
function rootFolder_() {
  const props = PropertiesService.getScriptProperties()
  const id = props.getProperty('ROOT_FOLDER_ID')
  if (id) {
    try {
      const f = DriveApp.getFolderById(id)
      if (!f.isTrashed()) return f
    } catch (err) { /* 消えていたら作り直す */ }
  }
  const folder = DriveApp.createFolder(ROOT_FOLDER_NAME)
  props.setProperty('ROOT_FOLDER_ID', folder.getId())
  return folder
}

function testFolder_(testId, create) {
  const name = 'test_' + String(testId).replace(/[^\w-]/g, '')
  const it = rootFolder_().getFoldersByName(name)
  if (it.hasNext()) return it.next()
  return create ? rootFolder_().createFolder(name) : null
}

function fileIn_(folder, name) {
  const it = folder.getFilesByName(name)
  return it.hasNext() ? it.next() : null
}

/** すべてのテストの結果の一覧（新しい順）。読み込みは時間がかかるので、しばらく覚えておく */
function manifests_() {
  const cache = CacheService.getScriptCache()
  const hit = cache.get('manifests')
  if (hit) return JSON.parse(hit)
  const list = []
  const folders = rootFolder_().getFolders()
  while (folders.hasNext()) {
    const f = folders.next()
    const file = fileIn_(f, MANIFEST)
    if (!file) continue
    try { list.push(JSON.parse(file.getBlob().getDataAsString('UTF-8'))) } catch (err) { /* 壊れたファイルは飛ばす */ }
  }
  list.sort(function (a, b) { return String(b.createdAt).localeCompare(String(a.createdAt)) })
  const text = JSON.stringify(list)
  if (text.length < 90000) cache.put('manifests', text, 600)
  return list
}

function clearCache_() {
  CacheService.getScriptCache().remove('manifests')
}

// ---------- 先生用 ----------

/** 取り込みの始め：結果の一覧を保存する（同じテストは古いファイルを消して上書き） */
function adminBeginUpload(manifestJson) {
  assertTeacher_()
  const m = JSON.parse(manifestJson)
  if (!m || m.version !== 1 || !m.testId || !Array.isArray(m.students)) throw new Error('ダッシュボード用のZIPではないようです')
  const folder = testFolder_(m.testId, true)
  const files = folder.getFiles()
  while (files.hasNext()) files.next().setTrashed(true)
  folder.createFile(Utilities.newBlob(manifestJson, 'application/json', MANIFEST))
  clearCache_()
  return { title: m.title, students: m.students.length }
}

/** 答案画像を1枚ずつ保存する（base64） */
function adminUploadImage(testId, name, base64) {
  assertTeacher_()
  if (!/^images\/\d{3}\.jpg$/.test(name)) throw new Error('画像の名前が正しくありません')
  const folder = testFolder_(testId, false)
  if (!folder) throw new Error('先に結果の一覧を取り込んでください')
  const fileName = name.replace('images/', '')
  const old = fileIn_(folder, fileName)
  if (old) old.setTrashed(true)
  folder.createFile(Utilities.newBlob(Utilities.base64Decode(base64), 'image/jpeg', fileName))
  return true
}

/** 取り込み済みのテストの一覧 */
function adminListTests() {
  assertTeacher_()
  clearCache_()
  return manifests_().map(function (m) {
    return {
      testId: m.testId, title: m.title, createdAt: m.createdAt, count: m.count, students: m.students.length,
      average: m.average, max: m.max,
      emails: m.students.map(function (s) { return { email: s.email, label: s.label } }),
    }
  })
}

function adminDeleteTest(testId) {
  assertTeacher_()
  const folder = testFolder_(testId, false)
  if (folder) folder.setTrashed(true)
  clearCache_()
  return true
}

// ---------- 生徒用 ----------

/** 結果を見る人：ふつうはログインしている本人。先生だけは確認のため別の生徒を指定できる */
function viewer_(previewAs) {
  if (previewAs && isTeacher_()) return String(previewAs).trim().toLowerCase()
  return me_()
}

/** 本人の結果（ほかの生徒の情報は返さない） */
function getMyResults(previewAs) {
  const email = viewer_(previewAs)
  const out = []
  manifests_().forEach(function (m) {
    const s = m.students.filter(function (x) { return x.email === email })[0]
    if (!s) return
    out.push({
      testId: m.testId, title: m.title, createdAt: m.createdAt,
      label: s.label, total: s.total, knowledge: s.knowledge, thinking: s.thinking, marks: s.marks,
      questions: m.questions, max: m.max, maxKnowledge: m.maxKnowledge, maxThinking: m.maxThinking,
      count: m.count, average: m.average, distribution: m.distribution, bin: s.bin,
      rates: rates_(m),
    })
  })
  return { email: email, results: out }
}

/**
 * 問題ごとの正答率（〇の人数 ÷ その問題を採点した人数。部分点は正答に数えない）。採点した人がいなければ null。
 * 学年・クラス全体の割合だけを返し、ほかの生徒の採点結果そのものは返さない
 */
function rates_(m) {
  return m.questions.map(function (q, k) {
    var n = 0, ok = 0
    m.students.forEach(function (s) {
      var v = s.marks[k]
      if (v === null || v === undefined) return
      n++
      if (v === 'o') ok++
    })
    return n ? ok / n : null
  })
}

/** 本人の答案画像（data URL） */
function getMyImage(testId, previewAs) {
  const email = viewer_(previewAs)
  const m = manifests_().filter(function (x) { return x.testId === testId })[0]
  if (!m) throw new Error('テストが見つかりません')
  const s = m.students.filter(function (x) { return x.email === email })[0]
  if (!s) throw new Error('このテストの結果はありません')
  const folder = testFolder_(testId, false)
  const file = folder && fileIn_(folder, s.image.replace('images/', ''))
  if (!file) throw new Error('答案の画像が見つかりません')
  return 'data:image/jpeg;base64,' + Utilities.base64Encode(file.getBlob().getBytes())
}
