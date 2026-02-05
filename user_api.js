// =======================================
// user_api.gs (権限フィルター版)
// =======================================

/**
 * ログインユーザー情報を返す API
 */
function getCurrentUserInfo() {
  var email = getActiveUserEmail_(); // ログイン中のメールアドレス
  var user = findUserByEmail_(email); // マスタから検索

  var name = user && user.name ? user.name : email;
  var role = user && user.role ? user.role : 'APPLICANT';
  var approverEmail = user && user.managerEmail ? user.managerEmail : '';

  return {
    id: user ? user.id : '',
    email: email,
    name: name,
    role: role,
    approverEmail: approverEmail
  };
}

/**
 * 初期表示用データ取得
 */
function getInitData() {
  // まず自分の情報を取得
  var currentUser = getCurrentUserInfo();

  return {
    currentUser: currentUser,
    // 自分の情報を渡して、リストをフィルタリングしてもらう
    list: getApplicantList(currentUser),
    // 科目マスタ（税率情報付き）も取得
    subjectMaster: getSubjectMasterWithTax_(),
    // 拠点リストも取得
    branches: getBranchList_()
  };
}

/**
 * 権限に応じてユーザーリストを絞り込む
 */
function getApplicantList(requestingUser) {
  // 引数がない場合の保険
  if (!requestingUser) requestingUser = getCurrentUserInfo();

  var sheet = getSheet_(SHEET_NAMES.USER_MASTER);
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];

  // A列(ID) 〜 F列(拠点) までを取得
  var values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  var list = [];
  var isAdmin = (requestingUser.role === 'ADMIN');
  var myEmail = normalizeEmail_(requestingUser.email);

  for (var i = 0; i < values.length; i++) {
    var id = values[i][0];    // A列
    var email = values[i][1]; // B列
    var name = values[i][2];  // C列
    var rowEmail = normalizeEmail_(email);

    // ★ここがフィルターロジック
    // 「管理者である」または「自分のメールアドレスと一致する」場合のみリストに入れる
    if (isAdmin || rowEmail === myEmail) {
      if (id && name) {
        list.push({
          id: String(id),
          email: rowEmail,
          name: String(name),
          branch: values[i][5] || '' // ★追加: 拠点
        });
      }
    }
  }
  return list;
}

/**
 * IDでユーザー検索 (内部用)
 */
function findUserById_(targetId) {
  var values = getUserMasterData_();
  if (!values) return null;

  var target = String(targetId);

  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === target) {
      return mapUserRow_(values[i]);
    }
  }
  return null;
}

/**
 * メールアドレスでユーザー検索 (内部用)
 */
function findUserByEmail_(email) {
  var values = getUserMasterData_();
  if (!values) return null;

  var target = normalizeEmail_(email);

  for (var i = 0; i < values.length; i++) {
    var rowEmail = normalizeEmail_(values[i][1]);
    if (rowEmail === target) {
      return mapUserRow_(values[i]);
    }
  }
  return null;
}

/**
 * 共通: ユーザーマスタの全データ取得
 */
function getUserMasterData_() {
  var ss = getAppSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.USER_MASTER);
  if (!sheet) return null;
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;

  return sheet.getRange(2, 1, lastRow - 1, 6).getValues();
}

/**
 * 共通: 行データをオブジェクトに変換
 */
function mapUserRow_(row) {
  return {
    id: row[0],
    email: row[1],
    name: row[2],
    role: String(row[3] || '').trim().toUpperCase(), // ★正規化: 大文字に統一
    managerEmail: row[4],
    branch: row[5] || '' // ★追加: 拠点
  };
}