const { google } = require('googleapis');

exports.handler = async (event) => {
  // CORS 헤더
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Content-Type': 'application/json',
  };

  // OPTIONS 요청 처리 (preflight)
  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers, body: '' };
  }

  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, headers, body: JSON.stringify({ error: 'Method not allowed' }) };
  }

  try {
    const data = JSON.parse(event.body);

    // Google Auth
    const auth = new google.auth.GoogleAuth({
      credentials: {
        client_email: process.env.GOOGLE_CLIENT_EMAIL,
        private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
      },
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const sheets = google.sheets({ version: 'v4', auth });
    const spreadsheetId = process.env.GOOGLE_SHEET_ID;

    // 공고목록 저장
    if (data.jobs && data.jobs.length > 0) {
      const jobRows = [
        ['공고순서','공고명','카테고리','부서','경력조건','근무지','상태','Tally_URL','마감일','노출여부'],
        ...data.jobs.map((j, i) => [
          i + 1, j.title||'', j.category||'', j.dept||'',
          j.career||'', j.location||'', j.status||'채용중',
          j.tallyUrl||'', j.deadline||'채용시마감', j.show||'Y',
        ])
      ];

      // 기존 데이터 지우고 새로 쓰기
      await sheets.spreadsheets.values.clear({
        spreadsheetId,
        range: '공고목록!A:J',
      });
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: '공고목록!A1',
        valueInputOption: 'RAW',
        requestBody: { values: jobRows },
      });
    }

    // 공고상세 저장
    if (data.details) {
      // 기존 상세 데이터 읽기
      let existingDetails = {};
      try {
        const res = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: '공고상세!A:B',
        });
        const rows = res.data.values || [];
        rows.slice(1).forEach(row => {
          if (row[0]) {
            try { existingDetails[row[0]] = JSON.parse(row[1] || '{}'); }
            catch(e) { existingDetails[row[0]] = {}; }
          }
        });
      } catch(e) {}

      // 새 데이터 병합
      Object.assign(existingDetails, data.details);

      // 공고상세 시트가 없으면 먼저 생성
      try {
        await sheets.spreadsheets.values.clear({
          spreadsheetId,
          range: '공고상세!A:B',
        });
      } catch(e) {
        // 시트 없으면 생성
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: {
            requests: [{ addSheet: { properties: { title: '공고상세' } } }]
          }
        });
      }

      const detailRows = [
        ['공고명', '상세내용(JSON)'],
        ...Object.entries(existingDetails).map(([title, detail]) => [
          title, JSON.stringify(detail)
        ])
      ];

      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: '공고상세!A1',
        valueInputOption: 'RAW',
        requestBody: { values: detailRows },
      });
    }

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify({ success: true }),
    };

  } catch (err) {
    console.error('Error:', err);
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ error: err.message }),
    };
  }
};
