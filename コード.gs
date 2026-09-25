/*************************************************
 * アンスーリール・ド・サクラ
 * 運営指導・監査管理センター
 *
 * Code.gs
 * 書類管理強化 + 職員管理 完全版
 *************************************************/


const CONFIG = {

  WEB_MASTER_SHEET: '10_Web連携マスター',

  USER_SHEET: '11_ログインユーザー',

  LOGIN_LOG_SHEET: '12_ログイン履歴',

  ADMIN_LOG_SHEET: '13_管理操作履歴',

  SOURCE_SHEET: '07_完全版チェックリスト',


  SESSION_SECONDS: 21600,

  SESSION_PREFIX: 'ANSOURIRE_SESSION_'

};


/* =================================================
 * Webアプリ
 * ================================================= */

function doGet() {

  return HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setTitle('運営指導・監査管理センター')
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1'
    );

}


/* =================================================
 * ログイン
 * ================================================= */

function loginUser(loginId, password) {

  loginId =
    String(loginId || '').trim();

  password =
    String(password || '');


  if (!loginId || !password) {

    return {
      success: false,
      message: 'IDとパスワードを入力してください。'
    };

  }


  ensureSystemSheets_();


  const user =
    findUserByLoginId_(loginId);


  if (!user) {

    recordLogin_(
      loginId,
      '',
      false,
      'ID不存在'
    );

    return {
      success: false,
      message: 'IDまたはパスワードが違います。'
    };

  }


  if (user.status !== '有効') {

    recordLogin_(
      loginId,
      user.name,
      false,
      '停止中アカウント'
    );

    return {
      success: false,
      message: 'このアカウントは現在利用できません。'
    };

  }


  const inputHash =
    hashPassword(password);


  if (user.passwordHash !== inputHash) {

    recordLogin_(
      loginId,
      user.name,
      false,
      'パスワード不一致'
    );

    return {
      success: false,
      message: 'IDまたはパスワードが違います。'
    };

  }


  const token =
    createSessionToken_();


  saveSession_(
    token,
    {
      loginId: user.loginId,
      name: user.name
    }
  );


  recordLogin_(
    loginId,
    user.name,
    true,
    'ログイン成功'
  );


  return {

    success: true,

    token: token,

    name: user.name,

    role: user.role,

    expiresIn:
      CONFIG.SESSION_SECONDS

  };

}


/* =================================================
 * セッション
 * ================================================= */

function createSessionToken_() {

  return (
    Utilities.getUuid() +
    Utilities.getUuid() +
    String(Date.now())
  ).replace(/-/g, '');

}


function saveSession_(token, user) {

  const cache =
    CacheService.getScriptCache();


  const data = {

    loginId:
      user.loginId,

    name:
      user.name,

    createdAt:
      new Date().toISOString()

  };


  cache.put(

    CONFIG.SESSION_PREFIX + token,

    JSON.stringify(data),

    CONFIG.SESSION_SECONDS

  );

}


function getSession_(token) {

  token =
    String(token || '').trim();


  if (!token) {

    return null;

  }


  const cache =
    CacheService.getScriptCache();


  const key =
    CONFIG.SESSION_PREFIX + token;


  const raw =
    cache.get(key);


  if (!raw) {

    return null;

  }


  try {

    const session =
      JSON.parse(raw);


    const currentUser =
      findUserByLoginId_(
        session.loginId
      );


    if (
      !currentUser ||
      currentUser.status !== '有効'
    ) {

      cache.remove(key);

      return null;

    }


    session.name =
      currentUser.name;

    session.role =
      currentUser.role;


    cache.put(

      key,

      JSON.stringify({

        loginId:
          currentUser.loginId,

        name:
          currentUser.name,

        createdAt:
          session.createdAt || ''

      }),

      CONFIG.SESSION_SECONDS

    );


    return session;


  } catch (error) {

    cache.remove(key);

    return null;

  }

}


function validateSession(token) {

  const session =
    getSession_(token);


  if (!session) {

    return {
      valid: false
    };

  }


  return {

    valid: true,

    loginId:
      session.loginId,

    name:
      session.name,

    role:
      session.role

  };

}


function logoutUser(token) {

  token =
    String(token || '').trim();


  if (!token) {

    return true;

  }


  CacheService
    .getScriptCache()
    .remove(
      CONFIG.SESSION_PREFIX + token
    );


  return true;

}


function requireSession_(token) {

  const session =
    getSession_(token);


  if (!session) {

    throw new Error(
      '認証の有効期限が切れています。再度ログインしてください。'
    );

  }


  return session;

}


function requireAdminSession_(token) {

  const session =
    requireSession_(token);


  const user =
    findUserByLoginId_(
      session.loginId
    );


  if (
    !user ||
    user.status !== '有効' ||
    user.role !== '管理者'
  ) {

    throw new Error(
      'この操作を行う権限がありません。'
    );

  }


  return user;

}


/* =================================================
 * ユーザー検索
 * ================================================= */

function findUserByLoginId_(loginId) {

  ensureSystemSheets_();


  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName(
      CONFIG.USER_SHEET
    );


  if (!sheet) {

    return null;

  }


  const values =
    sheet
      .getDataRange()
      .getValues();


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const id =
      String(
        values[i][1] || ''
      ).trim();


    if (id === loginId) {

      return {

        row:
          i + 1,

        status:
          String(
            values[i][0] || ''
          ).trim(),

        loginId:
          id,

        name:
          String(
            values[i][2] || ''
          ).trim(),

        passwordHash:
          String(
            values[i][3] || ''
          ).trim(),

        role:
          String(
            values[i][4] || '一般'
          ).trim(),

        registeredAt:
          values[i][5],

        note:
          String(
            values[i][6] || ''
          ).trim()

      };

    }

  }


  return null;

}


/* =================================================
 * 書類一覧
 * ================================================= */

function getDocuments(token) {

  requireSession_(token);


  /*
   * 07_完全版チェックリストを正本とする。
   * Web表示の直前に必ず10_Web連携マスターを再構築するため、
   * 07で変更した状態・保管場所・URL・確認日がWebへ反映される。
   */
  rebuildWebMaster();


  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName(
      CONFIG.WEB_MASTER_SHEET
    );


  if (!sheet) {

    throw new Error(
      '10_Web連携マスターが見つかりません。'
    );

  }


  const values =
    sheet
      .getDataRange()
      .getDisplayValues();


  if (values.length < 2) {

    return [];

  }


  const headers =
    values[0];


  return values
    .slice(1)
    .filter(
      row =>
        String(
          row[0] || ''
        ).trim() === '表示'
    )
    .map(
      row => {

        const obj = {};


        headers.forEach(
          function(header, index) {

            obj[header] =
              row[index] || '';

          }
        );


        return obj;

      }
    );

}


/* =================================================
 * パスワード
 * ================================================= */

function hashPassword(password) {

  const digest =
    Utilities.computeDigest(

      Utilities.DigestAlgorithm.SHA_256,

      String(password),

      Utilities.Charset.UTF_8

    );


  return digest
    .map(
      function(byte) {

        const value =
          byte < 0
            ? byte + 256
            : byte;


        return (
          '0' +
          value.toString(16)
        ).slice(-2);

      }
    )
    .join('');

}


/* =================================================
 * システムシート
 * ================================================= */

function ensureSystemSheets_() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  let users =
    ss.getSheetByName(
      CONFIG.USER_SHEET
    );


  if (!users) {

    users =
      ss.insertSheet(
        CONFIG.USER_SHEET
      );


    users
      .getRange(
        1,
        1,
        1,
        7
      )
      .setValues([[
        '状態',
        'ログインID',
        '氏名',
        'パスワードハッシュ',
        '権限',
        '登録日',
        '備考'
      ]]);


    formatHeader_(
      users,
      7
    );


    const statusRule =
      SpreadsheetApp
        .newDataValidation()
        .requireValueInList(
          [
            '有効',
            '停止'
          ],
          true
        )
        .build();


    users
      .getRange(
        'A2:A1000'
      )
      .setDataValidation(
        statusRule
      );


    const roleRule =
      SpreadsheetApp
        .newDataValidation()
        .requireValueInList(
          [
            '管理者',
            '一般'
          ],
          true
        )
        .build();


    users
      .getRange(
        'E2:E1000'
      )
      .setDataValidation(
        roleRule
      );

  }


  let loginLogs =
    ss.getSheetByName(
      CONFIG.LOGIN_LOG_SHEET
    );


  if (!loginLogs) {

    loginLogs =
      ss.insertSheet(
        CONFIG.LOGIN_LOG_SHEET
      );


    loginLogs
      .getRange(
        1,
        1,
        1,
        5
      )
      .setValues([[
        '日時',
        'ログインID',
        '氏名',
        '結果',
        '備考'
      ]]);


    formatHeader_(
      loginLogs,
      5
    );

  }


  let adminLogs =
    ss.getSheetByName(
      CONFIG.ADMIN_LOG_SHEET
    );


  if (!adminLogs) {

    adminLogs =
      ss.insertSheet(
        CONFIG.ADMIN_LOG_SHEET
      );


    adminLogs
      .getRange(
        1,
        1,
        1,
        7
      )
      .setValues([[
        '日時',
        '操作者ID',
        '操作者氏名',
        '操作',
        '対象ID',
        '対象氏名',
        '備考'
      ]]);


    formatHeader_(
      adminLogs,
      7
    );

  }


  return {

    users:
      users,

    loginLogs:
      loginLogs,

    adminLogs:
      adminLogs

  };

}


function formatHeader_(
  sheet,
  columnCount
) {

  sheet
    .getRange(
      1,
      1,
      1,
      columnCount
    )
    .setBackground(
      '#0B5A9D'
    )
    .setFontColor(
      '#FFFFFF'
    )
    .setFontWeight(
      'bold'
    );


  sheet.setFrozenRows(1);

}


function createLoginSheets() {

  ensureSystemSheets_();

  SpreadsheetApp.flush();

  return true;

}


/* =================================================
 * ログ
 * ================================================= */

function recordLogin_(
  loginId,
  name,
  success,
  note
) {

  const sheets =
    ensureSystemSheets_();


  sheets.loginLogs.appendRow([

    new Date(),

    loginId,

    name,

    success
      ? '成功'
      : '失敗',

    note || ''

  ]);

}


function recordAdminAction_(
  admin,
  action,
  targetId,
  targetName,
  note
) {

  const sheets =
    ensureSystemSheets_();


  sheets.adminLogs.appendRow([

    new Date(),

    admin.loginId,

    admin.name,

    action,

    targetId || '',

    targetName || '',

    note || ''

  ]);

}


/* =================================================
 * 職員管理
 * ================================================= */

function getStaffUsers(token) {

  requireAdminSession_(token);


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.USER_SHEET
      );


  const values =
    sheet
      .getDataRange()
      .getDisplayValues();


  if (values.length < 2) {

    return [];

  }


  return values
    .slice(1)
    .filter(
      row =>
        String(
          row[1] || ''
        ).trim() !== ''
    )
    .map(
      row => ({

        status:
          String(
            row[0] || ''
          ).trim(),

        loginId:
          String(
            row[1] || ''
          ).trim(),

        name:
          String(
            row[2] || ''
          ).trim(),

        role:
          String(
            row[4] || '一般'
          ).trim(),

        registeredAt:
          String(
            row[5] || ''
          ).trim(),

        note:
          String(
            row[6] || ''
          ).trim()

      })
    );

}


function adminAddUser(
  token,
  loginId,
  name,
  password,
  role
) {

  const admin =
    requireAdminSession_(token);


  loginId =
    normalizeLoginId_(
      loginId
    );


  name =
    String(
      name || ''
    ).trim();


  password =
    String(
      password || ''
    );


  role =
    normalizeRole_(
      role
    );


  if (!loginId) {

    throw new Error(
      'ログインIDを入力してください。'
    );

  }


  if (!name) {

    throw new Error(
      '氏名を入力してください。'
    );

  }


  validatePassword_(
    password
  );


  if (
    findUserByLoginId_(
      loginId
    )
  ) {

    throw new Error(
      '同じログインIDがすでに登録されています。'
    );

  }


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.USER_SHEET
      );


  sheet.appendRow([

    '有効',

    loginId,

    name,

    hashPassword(password),

    role,

    new Date(),

    ''

  ]);


  SpreadsheetApp.flush();


  recordAdminAction_(

    admin,

    '職員追加',

    loginId,

    name,

    '権限：' + role

  );


  return {

    success: true,

    message:
      name +
      'さんを登録しました。'

  };

}


function adminSetUserStatus(
  token,
  targetLoginId,
  newStatus
) {

  const admin =
    requireAdminSession_(token);


  targetLoginId =
    String(
      targetLoginId || ''
    ).trim();


  newStatus =
    String(
      newStatus || ''
    ).trim();


  if (
    newStatus !== '有効' &&
    newStatus !== '停止'
  ) {

    throw new Error(
      '状態の指定が正しくありません。'
    );

  }


  const target =
    findUserByLoginId_(
      targetLoginId
    );


  if (!target) {

    throw new Error(
      '対象職員が見つかりません。'
    );

  }


  if (
    target.loginId ===
      admin.loginId &&
    newStatus === '停止'
  ) {

    throw new Error(
      '現在ログイン中の自分自身を停止することはできません。'
    );

  }


  if (
    target.role === '管理者' &&
    target.status === '有効' &&
    newStatus === '停止'
  ) {

    ensureAnotherActiveAdmin_(
      target.loginId
    );

  }


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.USER_SHEET
      );


  sheet
    .getRange(
      target.row,
      1
    )
    .setValue(
      newStatus
    );


  SpreadsheetApp.flush();


  recordAdminAction_(

    admin,

    newStatus === '有効'
      ? 'アカウント有効化'
      : 'アカウント停止',

    target.loginId,

    target.name,

    ''

  );


  return {

    success: true,

    message:
      target.name +
      'さんを' +
      newStatus +
      'にしました。'

  };

}


function adminSetUserRole(
  token,
  targetLoginId,
  newRole
) {

  const admin =
    requireAdminSession_(token);


  targetLoginId =
    String(
      targetLoginId || ''
    ).trim();


  newRole =
    normalizeRole_(
      newRole
    );


  const target =
    findUserByLoginId_(
      targetLoginId
    );


  if (!target) {

    throw new Error(
      '対象職員が見つかりません。'
    );

  }


  if (
    target.loginId ===
      admin.loginId &&
    newRole !== '管理者'
  ) {

    throw new Error(
      '現在ログイン中の自分自身を一般職員へ変更することはできません。'
    );

  }


  if (
    target.role === '管理者' &&
    target.status === '有効' &&
    newRole !== '管理者'
  ) {

    ensureAnotherActiveAdmin_(
      target.loginId
    );

  }


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.USER_SHEET
      );


  sheet
    .getRange(
      target.row,
      5
    )
    .setValue(
      newRole
    );


  SpreadsheetApp.flush();


  recordAdminAction_(

    admin,

    '権限変更',

    target.loginId,

    target.name,

    target.role +
      ' → ' +
      newRole

  );


  return {

    success: true,

    message:
      target.name +
      'さんの権限を' +
      newRole +
      'に変更しました。'

  };

}


function adminResetPassword(
  token,
  targetLoginId,
  newPassword
) {

  const admin =
    requireAdminSession_(token);


  targetLoginId =
    String(
      targetLoginId || ''
    ).trim();


  newPassword =
    String(
      newPassword || ''
    );


  validatePassword_(
    newPassword
  );


  const target =
    findUserByLoginId_(
      targetLoginId
    );


  if (!target) {

    throw new Error(
      '対象職員が見つかりません。'
    );

  }


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.USER_SHEET
      );


  sheet
    .getRange(
      target.row,
      4
    )
    .setValue(
      hashPassword(
        newPassword
      )
    );


  const changedAt =
    Utilities.formatDate(

      new Date(),

      Session.getScriptTimeZone(),

      'yyyy/MM/dd HH:mm'

    );


  sheet
    .getRange(
      target.row,
      7
    )
    .setValue(
      'パスワード再設定：' +
      changedAt
    );


  SpreadsheetApp.flush();


  recordAdminAction_(

    admin,

    'パスワード再設定',

    target.loginId,

    target.name,

    'パスワード本文は記録していません。'

  );


  return {

    success: true,

    message:
      target.name +
      'さんのパスワードを再設定しました。'

  };

}


function ensureAnotherActiveAdmin_(
  excludeLoginId
) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.USER_SHEET
      );


  const values =
    sheet
      .getDataRange()
      .getValues();


  let count = 0;


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const status =
      String(
        values[i][0] || ''
      ).trim();


    const loginId =
      String(
        values[i][1] || ''
      ).trim();


    const role =
      String(
        values[i][4] || ''
      ).trim();


    if (
      loginId !== excludeLoginId &&
      status === '有効' &&
      role === '管理者'
    ) {

      count++;

    }

  }


  if (count < 1) {

    throw new Error(
      '有効な管理者が0人になるため、この操作はできません。'
    );

  }

}


function normalizeLoginId_(
  loginId
) {

  const value =
    String(
      loginId || ''
    )
      .trim()
      .toLowerCase();


  if (!value) {

    return '';

  }


  if (
    !/^[a-z0-9._-]+$/.test(
      value
    )
  ) {

    throw new Error(
      'ログインIDは半角英数字と「.」「_」「-」のみ使用できます。'
    );

  }


  if (
    value.length < 3 ||
    value.length > 40
  ) {

    throw new Error(
      'ログインIDは3〜40文字で設定してください。'
    );

  }


  return value;

}


function normalizeRole_(
  role
) {

  role =
    String(
      role || '一般'
    ).trim();


  if (
    role !== '管理者' &&
    role !== '一般'
  ) {

    throw new Error(
      '権限の指定が正しくありません。'
    );

  }


  return role;

}


function validatePassword_(
  password
) {

  password =
    String(
      password || ''
    );


  if (
    password.length < 8
  ) {

    throw new Error(
      '新しいパスワードは8文字以上で設定してください。'
    );

  }


  if (
    password.length > 100
  ) {

    throw new Error(
      'パスワードが長すぎます。'
    );

  }

}


/* =================================================
 * 書類管理強化
 *
 * 07_完全版チェックリスト
 *
 * A No.
 * B 区分
 * C 確認項目
 * D 必要書類・証跡
 * E 県資料上の根拠
 * F 主担当
 * G 確認頻度
 * H 重要度
 * I 状態
 * J 保管場所
 * K 直接リンク
 * L 最終確認日
 * M 備考
 *
 * N 次回確認期限
 * O 自動判定
 * P 確認ポイント
 * Q 更新担当メモ
 * R 更新日
 * ================================================= */


/* =================================================
 * 初回セットアップ
 *
 * ★ これを1回だけ手動実行
 * ================================================= */

function setupDocumentManagement() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName(
      CONFIG.SOURCE_SHEET
    );


  if (!sheet) {

    throw new Error(
      '07_完全版チェックリストが見つかりません。'
    );

  }


  /*
   * 既存A～Mには触れない。
   * N～Rだけ管理列として使用。
   */

  const headerRow =
    4;


  const headers = [

    '次回確認期限',

    '自動判定',

    '確認ポイント',

    '更新担当メモ',

    '更新日'

  ];


  sheet
    .getRange(
      headerRow,
      14,
      1,
      headers.length
    )
    .setValues([
      headers
    ]);


  sheet
    .getRange(
      headerRow,
      14,
      1,
      headers.length
    )
    .setBackground(
      '#0B5A9D'
    )
    .setFontColor(
      '#FFFFFF'
    )
    .setFontWeight(
      'bold'
    )
    .setHorizontalAlignment(
      'center'
    );


  /*
   * I列「状態」の候補を統一
   */

  const lastRow =
    Math.max(
      sheet.getLastRow(),
      5
    );


  const statusRule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        [
          '未確認',
          '確認済',
          '整備済',
          '要更新',
          '不足'
        ],
        true
      )
      .setAllowInvalid(true)
      .build();


  sheet
    .getRange(
      5,
      9,
      Math.max(
        lastRow - 4,
        1
      ),
      1
    )
    .setDataValidation(
      statusRule
    );


  /*
   * N・Rは日付
   */

  sheet
    .getRange(
      5,
      14,
      Math.max(
        lastRow - 4,
        1
      ),
      1
    )
    .setNumberFormat(
      'yyyy/mm/dd'
    );


  sheet
    .getRange(
      5,
      18,
      Math.max(
        lastRow - 4,
        1
      ),
      1
    )
    .setNumberFormat(
      'yyyy/mm/dd'
    );


  /*
   * 新しい列だけ幅調整
   */

  sheet.setColumnWidth(
    14,
    120
  );

  sheet.setColumnWidth(
    15,
    110
  );

  sheet.setColumnWidth(
    16,
    280
  );

  sheet.setColumnWidth(
    17,
    280
  );

  sheet.setColumnWidth(
    18,
    110
  );


  /*
   * 既存89件を読み取り、
   * N/Oだけ計算。
   *
   * P/Q/Rの既存値は消さない。
   */

  refreshDocumentManagement_();


  /*
   * Web用データも再構築
   */

  const result =
    rebuildWebMaster();


  SpreadsheetApp.flush();


  return {

    success: true,

    message:
      '書類管理機能の初期設定が完了しました。',

    count:
      result.count

  };

}


/* =================================================
 * 管理情報再計算
 *
 * N 次回期限
 * O 自動判定
 *
 * P/Q/Rは触らない
 * ================================================= */

function refreshDocumentManagement() {

  refreshDocumentManagement_();

  const result =
    rebuildWebMaster();


  SpreadsheetApp.flush();


  return {

    success: true,

    message:
      '書類管理情報を更新しました。',

    count:
      result.count

  };

}


function refreshDocumentManagement_() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName(
      CONFIG.SOURCE_SHEET
    );


  if (!sheet) {

    throw new Error(
      '07_完全版チェックリストが見つかりません。'
    );

  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 5) {

    return;

  }


  const rowCount =
    lastRow - 4;


  /*
   * A～Rまで取得
   */

  const values =
    sheet
      .getRange(
        5,
        1,
        rowCount,
        18
      )
      .getValues();


  const nextDates = [];

  const autoStatuses = [];


  values.forEach(
    function(row) {

      const no =
        row[0];


      if (!no) {

        nextDates.push([
          ''
        ]);

        autoStatuses.push([
          ''
        ]);

        return;

      }


      const frequency =
        String(
          row[6] || ''
        ).trim();


      const manualStatus =
        String(
          row[8] || ''
        ).trim();


      const storage =
        String(
          row[9] || ''
        ).trim();


      const url =
        String(
          row[10] || ''
        ).trim();


      const lastChecked =
        parseDate_(
          row[11]
        );


      const nextDate =
        calculateNextReviewDate_(
          frequency,
          lastChecked
        );


      const autoStatus =
        calculateAutoStatus_(

          manualStatus,

          storage,

          url,

          lastChecked,

          nextDate,

          frequency

        );


      nextDates.push([
        nextDate || ''
      ]);


      autoStatuses.push([
        autoStatus
      ]);

    }
  );


  sheet
    .getRange(
      5,
      14,
      rowCount,
      1
    )
    .setValues(
      nextDates
    )
    .setNumberFormat(
      'yyyy/mm/dd'
    );


  sheet
    .getRange(
      5,
      15,
      rowCount,
      1
    )
    .setValues(
      autoStatuses
    );


  /*
   * 自動判定の色分け
   */

  applyAutoStatusFormatting_(
    sheet,
    rowCount
  );

}


/* =================================================
 * 次回確認期限
 * ================================================= */

function calculateNextReviewDate_(
  frequency,
  lastChecked
) {

  if (!lastChecked) {

    return null;

  }


  const text =
    String(
      frequency || ''
    ).trim();


  if (!text) {

    return null;

  }


  /*
   * 随時・実施時などは固定期限なし
   */

  if (
    /随時|実施時|必要時|都度|変更時|発生時/.test(
      text
    )
  ) {

    return null;

  }


  const date =
    new Date(
      lastChecked.getFullYear(),
      lastChecked.getMonth(),
      lastChecked.getDate()
    );


  if (
    /毎月|月次|月1|1か月|１か月/.test(
      text
    )
  ) {

    date.setMonth(
      date.getMonth() + 1
    );

    return date;

  }


  if (
    /3か月|３か月|四半期|年4回|年４回/.test(
      text
    )
  ) {

    date.setMonth(
      date.getMonth() + 3
    );

    return date;

  }


  if (
    /半年|6か月|６か月|年2回|年２回/.test(
      text
    )
  ) {

    date.setMonth(
      date.getMonth() + 6
    );

    return date;

  }


  if (
    /年次|毎年|年1回|年１回|1年|１年/.test(
      text
    )
  ) {

    date.setFullYear(
      date.getFullYear() + 1
    );

    return date;

  }


  /*
   * 頻度を安全に判断できないものは
   * 勝手に期限を作らない
   */

  return null;

}


/* =================================================
 * 自動判定
 * ================================================= */

function calculateAutoStatus_(
  manualStatus,
  storage,
  url,
  lastChecked,
  nextDate,
  frequency
) {

  const manual =
    String(
      manualStatus || ''
    ).trim();


  /*
   * 人が「不足」と確認したものは最優先
   */

  if (
    manual === '不足' ||
    manual.includes('不足')
  ) {

    return '不足';

  }


  /*
   * 人が要更新と判断
   */

  if (
    manual === '要更新' ||
    manual.includes('要更新')
  ) {

    return '要更新';

  }


  /*
   * 期限超過
   */

  if (nextDate) {

    const today =
      startOfDay_(
        new Date()
      );


    const due =
      startOfDay_(
        nextDate
      );


    if (
      due.getTime() <
      today.getTime()
    ) {

      return '要更新';

    }


    /*
     * 30日以内に期限到来
     */

    const diffDays =
      Math.ceil(
        (
          due.getTime() -
          today.getTime()
        ) /
        86400000
      );


    if (
      diffDays >= 0 &&
      diffDays <= 30
    ) {

      return '要確認';

    }

  }


  /*
   * 保管場所・リンクの双方がない場合は
   * 不足と断定せず棚卸し対象
   */

  if (
    !storage &&
    !url
  ) {

    return '要確認';

  }


  /*
   * 最終確認日が未入力
   */

  if (!lastChecked) {

    return '要確認';

  }


  /*
   * 人が確認済・整備済としたもの
   */

  if (
    manual === '確認済' ||
    manual === '整備済'
  ) {

    return '整備済';

  }


  /*
   * 随時書類でも
   * 所在確認＋最終確認済なら整備扱い
   */

  if (
    /随時|実施時|必要時|都度|変更時|発生時/.test(
      String(
        frequency || ''
      )
    ) &&
    (
      storage ||
      url
    ) &&
    lastChecked
  ) {

    return '整備済';

  }


  return '要確認';

}


/* =================================================
 * 日付補助
 * ================================================= */

function parseDate_(value) {

  if (!value) {

    return null;

  }


  if (
    Object.prototype.toString.call(value) ===
      '[object Date]' &&
    !isNaN(
      value.getTime()
    )
  ) {

    return new Date(
      value.getFullYear(),
      value.getMonth(),
      value.getDate()
    );

  }


  const text =
    String(
      value
    ).trim();


  if (!text) {

    return null;

  }


  const normalized =
    text.replace(
      /[.\-]/g,
      '/'
    );


  const parts =
    normalized.split('/');


  if (
    parts.length === 3
  ) {

    const y =
      Number(
        parts[0]
      );


    const m =
      Number(
        parts[1]
      );


    const d =
      Number(
        parts[2]
      );


    if (
      y &&
      m &&
      d
    ) {

      const date =
        new Date(
          y,
          m - 1,
          d
        );


      if (
        !isNaN(
          date.getTime()
        )
      ) {

        return date;

      }

    }

  }


  const parsed =
    new Date(
      text
    );


  if (
    isNaN(
      parsed.getTime()
    )
  ) {

    return null;

  }


  return new Date(
    parsed.getFullYear(),
    parsed.getMonth(),
    parsed.getDate()
  );

}


function startOfDay_(date) {

  return new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate()
  );

}


/* =================================================
 * 自動判定 色分け
 * ================================================= */

function applyAutoStatusFormatting_(
  sheet,
  rowCount
) {

  if (
    rowCount < 1
  ) {

    return;

  }


  const range =
    sheet.getRange(
      5,
      15,
      rowCount,
      1
    );


  const rules =
    sheet
      .getConditionalFormatRules()
      .filter(
        function(rule) {

          const ranges =
            rule.getRanges();


          return !ranges.some(
            function(r) {

              return (
                r.getColumn() === 15 &&
                r.getSheet()
                  .getName() ===
                  sheet.getName()
              );

            }
          );

        }
      );


  rules.push(

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextEqualTo(
        '整備済'
      )
      .setBackground(
        '#EAF7EF'
      )
      .setFontColor(
        '#157347'
      )
      .setRanges([
        range
      ])
      .build(),

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextEqualTo(
        '要確認'
      )
      .setBackground(
        '#FFF7E8'
      )
      .setFontColor(
        '#9A5B00'
      )
      .setRanges([
        range
      ])
      .build(),

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextEqualTo(
        '要更新'
      )
      .setBackground(
        '#FFF0E6'
      )
      .setFontColor(
        '#B54708'
      )
      .setRanges([
        range
      ])
      .build(),

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextEqualTo(
        '不足'
      )
      .setBackground(
        '#FFF1F0'
      )
      .setFontColor(
        '#B42318'
      )
      .setRanges([
        range
      ])
      .build()

  );


  sheet.setConditionalFormatRules(
    rules
  );

}


/* =================================================
 * Webマスター再構築
 * ================================================= */

function syncChecklistToWebMaster() {

  const result =
    rebuildWebMaster();

  SpreadsheetApp
    .getActiveSpreadsheet()
    .toast(
      '07 → 10 の同期完了：' +
      result.count +
      '件',
      'Web連携',
      5
    );

  return result;

}


function rebuildWebMaster() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const source =
    ss.getSheetByName(
      CONFIG.SOURCE_SHEET
    );
if (!source) {

    throw new Error(
      '07_完全版チェックリストが見つかりません。'
    );

  }


  let master =
    ss.getSheetByName(
      CONFIG.WEB_MASTER_SHEET
    );


  if (!master) {

    master =
      ss.insertSheet(
        CONFIG.WEB_MASTER_SHEET
      );

  }


  /*
   * 10は派生データなので再構築可
   */

  master.clear();


  const headers = [

    'Web表示',

    '表示順',

    'カテゴリー',

    '書類名',

    '状態',

    '担当者',

    '重要度',

    '最終確認日',

    '次回確認期限',

    'Google Drive / 書類URL',

    '備考',

    '検索キーワード',

    '更新日時',

    'Web用ID',

    'リンク状態',

    '保管場所',

    '確認ポイント',

    '更新担当メモ',

    '確認頻度',

    '手動状態',

    '県資料上の根拠'

  ];


  master
    .getRange(
      1,
      1,
      1,
      headers.length
    )
    .setValues([
      headers
    ]);

  const lastRow =
    source.getLastRow();


  if (lastRow < 5) {

    throw new Error(
      '07_完全版チェックリストにデータがありません。'
    );

  }


  /*
   * A～R
   */

  const rawValues =
    source
      .getRange(
        5,
        1,
        lastRow - 4,
        18
      )
      .getValues();


  const displayValues =
    source
      .getRange(
        5,
        1,
        lastRow - 4,
        18
      )
      .getDisplayValues();


  /*
   * K列のDriveスマートチップ / ハイパーリンク対応
   * 表示文字ではなく実URLを取得するためRichTextValueも読む。
   */
  const richTextValues =
    source
      .getRange(
        5,
        1,
        lastRow - 4,
        18
      )
      .getRichTextValues();



  const rows = [];


  for (
    let i = 0;
    i < rawValues.length;
    i++
  ) {

    const raw =
      rawValues[i];


    const row =
      displayValues[i];


    const no =
      Number(
        raw[0]
      );


    if (!no) {

      continue;

    }


    const type =
      String(
        row[1] || ''
      ).trim();


    const item =
      String(
        row[2] || ''
      ).trim();


    const evidence =
      String(
        row[3] || ''
      ).trim();


    const prefecturalBasis =
      String(
        row[4] || ''
      ).trim();


    const owner =
      String(
        row[5] || ''
      ).trim();


    const frequency =
      String(
        row[6] || ''
      ).trim();


    const importance =
      String(
        row[7] || ''
      ).trim();


    const manualStatus =
      String(
        row[8] || ''
      ).trim();


    const storage =
      String(
        row[9] || ''
      ).trim();


    let url =
      extractDocumentUrl_(
        raw[10],
        row[10],
        richTextValues[i][10]
      );


    const lastCheck =
      String(
        row[11] || ''
      ).trim();


    const note =
      String(
        row[12] || ''
      ).trim();


    const nextCheck =
      String(
        row[13] || ''
      ).trim();


    const autoStatus =
      String(
        row[14] || ''
      ).trim();


    const checkPoint =
      String(
        row[15] || ''
      ).trim();


    const updateMemo =
      String(
        row[16] || ''
      ).trim();


    if (!item) {

      continue;

    }


    const category =
      determineWebCategory(
        type,
        item,
        evidence
      );

    const webStatus =
      determineWebStatus_(
        manualStatus,
        autoStatus,
        storage,
        url,
        lastCheck,
        nextCheck
      );


    rows.push([

      '表示',

      no,

      category,

      item,

      webStatus,

      owner,

      importance,

      lastCheck,

      nextCheck,

      url,

      note,

      [
        category,
        type,
        item,
        evidence,
        prefecturalBasis,
        owner,
        storage,
        checkPoint
      ]
        .filter(Boolean)
        .join(' '),

      new Date(),

      'DOC-' +
        String(no)
          .padStart(
            3,
            '0'
          ),

      url
        ? 'リンク登録済'
        : (
            storage
              ? '保管場所登録済'
              : '所在未登録'
          ),

      storage,

      checkPoint,

      updateMemo,

      frequency,

      manualStatus,

      prefecturalBasis

    ]);

  }


  if (rows.length) {

    master
      .getRange(
        2,
        1,
        rows.length,
        headers.length
      )
      .setValues(
        rows
      );

  }


  master.setFrozenRows(1);


  master
    .getRange(
      1,
      1,
      1,
      headers.length
    )
    .setBackground(
      '#0B5A9D'
    )
    .setFontColor(
      '#FFFFFF'
    )
    .setFontWeight(
      'bold'
    );


  master
    .getRange(
      1,
      1,
      Math.max(
        rows.length + 1,
        1
      ),
      headers.length
    )
    .setWrap(true)
    .setVerticalAlignment(
      'middle'
    );


  SpreadsheetApp.flush();


  return {

    success: true,

    count:
      rows.length

  };

}


/* =================================================
 * Web表示用ステータス判定
 * ================================================= */

function determineWebStatus_(
  manualStatus,
  autoStatus,
  storage,
  url,
  lastCheck,
  nextCheck
) {

  const manual =
    String(
      manualStatus || ''
    ).trim();

  const auto =
    String(
      autoStatus || ''
    ).trim();


  if (
    manual === '対象外'
  ) {

    return '対象外';

  }


  if (
    manual === '不足'
  ) {

    return '不足';

  }


  if (
    manual === '要更新'
  ) {

    return '要更新';

  }


  /*
   * 人が「確認済」としたものは、
   * 保管場所またはURLがあり、最終確認日もあれば確認済を優先。
   * O列の古い自動判定が「要確認」のままでもWeb表示を戻す。
   */
  if (
    /確認済|整備済|揃っている/.test(
      manual
    )
  ) {

    if (
      (storage || url) &&
      lastCheck
    ) {

      return '確認済';

    }

    return '要確認';

  }


  if (
    auto === '不足'
  ) {

    return '不足';

  }


  if (
    auto === '要更新'
  ) {

    return '要更新';

  }


  if (
    auto === '対象外'
  ) {

    return '対象外';

  }


  if (
    auto === '整備済' ||
    auto === '確認済'
  ) {

    return auto;

  }


  return '要確認';

}


/* =================================================
 * K列 URL取得
 * Driveスマートチップ / 通常ハイパーリンク / URL直貼り対応
 * ================================================= */

function extractDocumentUrl_(
  rawValue,
  displayValue,
  richTextValue
) {

  /*
   * 1. 通常のRichTextリンク
   */
  if (richTextValue) {

    const directUrl =
      richTextValue.getLinkUrl();

    if (
      directUrl &&
      /^https?:\/\//i.test(
        directUrl
      )
    ) {

      return directUrl;

    }


    /*
     * 一部のリンクはテキスト全体ではなく
     * Run単位にURLを持つため、Runも確認する。
     */
    const runs =
      richTextValue.getRuns();

    for (
      let i = 0;
      i < runs.length;
      i++
    ) {

      const runUrl =
        runs[i].getLinkUrl();

      if (
        runUrl &&
        /^https?:\/\//i.test(
          runUrl
        )
      ) {

        return runUrl;

      }

    }

  }


  /*
   * 2. URLを直接貼り付けている場合
   */
  const rawText =
    String(
      rawValue || ''
    ).trim();

  if (
    /^https?:\/\//i.test(
      rawText
    )
  ) {

    return rawText;

  }


  const displayText =
    String(
      displayValue || ''
    ).trim();

  if (
    /^https?:\/\//i.test(
      displayText
    )
  ) {

    return displayText;

  }


  /*
   * URLが取得できない場合は推測しない。
   * 無効なリンクをWebに出さないため空欄を返す。
   */
  return '';

}


/* =================================================
 * カテゴリー分類
 * ================================================= */

function determineWebCategory(
  type,
  item,
  evidence
) {

  const text =
    [
      type,
      item,
      evidence
    ]
      .filter(Boolean)
      .join(' ');


  if (
    /自主点検|点検調書/.test(
      text
    )
  ) {

    return '⑨ 自主点検';

  }


  if (
    /指定申請|指定更新|変更届|変更申請|体制等届|体制届|加算|基本報酬|処遇改善|業務管理体制|事業開始届|情報公表|WAM/.test(
      text
    )
  ) {

    return '⑧ 届出・加算';

  }


  if (
    /工賃|生産活動|施設外就労|施設外支援|就労支援|平均利用者|工賃向上|工賃引上/.test(
      text
    )
  ) {

    return '⑦ B型・工賃';

  }


  if (
    /請求|給付費|利用料|領収書|上限管理|法定代理|会計|決算|利用者負担額|サービス提供実績記録表/.test(
      text
    )
  ) {

    return '⑥ 請求・会計';

  }


  if (
    /虐待|身体拘束|職員研修|研修計画|研修記録/.test(
      text
    )
  ) {

    return '⑤ 虐待・身体拘束・研修';

  }


  if (
    /BCP|業務継続|災害|消防|避難|感染症|食中毒|衛生管理|非常災害/.test(
      text
    )
  ) {

    return '④ 防災・BCP・感染症';

  }


  if (
    /職員|勤務|組織|人事|雇用|給与|賃金台帳|就業規則|労基|健康診断|資格|履歴書|社会保険|雇用保険|労働条件|辞令|出勤簿|タイムカード|人員配置|ハラスメント/.test(
      text
    )
  ) {

    return '① 人員・労務';

  }


  if (
    /利用者|個別支援|支援記録|ケース記録|受給者証|サービス提供記録|利用契約|重要事項|地域交流|給食|預り金|モニタリング/.test(
      text
    )
  ) {

    return '② 利用者・支援';

  }


  return '③ 運営・設備';

}


/* =================================================
 * テスト
 * ================================================= */

function testSessionSystem() {

  const fakeToken =
    'THIS_IS_NOT_A_VALID_TOKEN';


  const result =
    validateSession(
      fakeToken
    );


  if (result.valid) {

    throw new Error(
      'テスト失敗：無効トークンが認証されました。'
    );

  }


  Logger.log(
    'OK：無効トークンは拒否されました。'
  );

}
