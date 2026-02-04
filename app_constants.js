// =======================================
// app_constants.gs
// =======================================

var TIMEZONE = 'Asia/Tokyo';

var SHEET_NAMES = {
  HEADER: '申請ヘッダ',
  DETAIL: '明細',
  SUBJECT_MASTER: '科目マスタ',
  USER_MASTER: 'ユーザーマスタ'
};

var ROLES = {
  APPLICANT: 'APPLICANT',
  APPROVER: 'APPROVER',
  ACCOUNTING: 'ACCOUNTING',
  ADMIN: 'ADMIN'
};

var STATUS = {
  DRAFT: '下書き',
  APPLYING: '申請中',
  APPROVED: '承認済',
  REJECTED: '却下',
  RETURNED: '差し戻し',
  FIXED: '経理確定'
};

var HEADER_COL = {
  APPLICATION_ID: 1,
  APPLICATION_DATE: 2,
  APPLICANT_EMAIL: 3,
  APPLICANT_NAME: 4,
  APPLICANT_DEPT: 5, // ★追加: 部署
  SERIAL_NO: 6,
  TOTAL_AMOUNT: 7,
  STATUS: 8,
  APPROVER_EMAIL: 9,
  APPROVED_AT: 10,
  REJECTED_AT: 11,
  RETURNED_AT: 12, // ★追加: 差し戻し日時
  RETURN_COMMENT: 13, // ★追加: 差し戻しコメント
  FIXED_AT: 14,
  NEEDS_CHECK: 15,
  REMARKS: 16,
  PAYMENT_STATUS: 17, // ★追加: 支払状況(未払い/支払済)
  COL_COUNT: 17
};

var DETAIL_COL = {
  DETAIL_ID: 1,
  APPLICATION_ID: 2,
  USAGE_DATE: 3,
  AMOUNT: 4,
  TAX_RATE: 5,
  TAX_AMOUNT: 6,
  VENDOR: 7,
  SUBJECT: 8,
  PAYMENT_METHOD: 9,
  PURPOSE: 10,
  RECEIPT_URL: 11,
  OCR_SCORE: 12,
  INVOICE_REG: 13,  // ★追加: インボイス登録有無
  COL_COUNT: 13
};

var USER_MASTER_COL = {
  ID: 1,
  EMAIL: 2,
  NAME: 3,
  ROLE: 4,
  MANAGER_EMAIL: 5,
  BRANCH: 6,  // ★追加: 拠点
  COL_COUNT: 6
};