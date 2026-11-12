/**
 * 학생 운동 기록 데이터베이스용 Google Apps Script
 * 대상 스프레드시트 ID: 1gilD5rjJ1rXxYix2hVeOW5MTVEQLg2nDkEo0hep6xLk
 * 
 * [시트 구성 안내]
 * 1. "기록" 탭: A열(이름), B열(종목), C열(기록), D열(시간) - 1행은 헤더
 * 2. "학생명단" 탭: A열(번호), B열(이름) - 1행은 헤더
 */

const SPREADSHEET_ID = '1gilD5rjJ1rXxYix2hVeOW5MTVEQLg2nDkEo0hep6xLk';
const RECORD_SHEET_NAME = '기록';
const STUDENT_SHEET_NAME = '학생명단';

/**
 * GET 요청 처리: "학생명단"과 최신순으로 정렬된 "기록" 데이터를 JSON으로 반환합니다.
 * 브라우저에서 일반 fetch(url)로 호출 시 CORS 문제 없이 JSON을 수신할 수 있습니다.
 */
function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 1. "기록" 시트 데이터 읽기 및 최신순 정렬
    const recordSheet = ss.getSheetByName(RECORD_SHEET_NAME);
    if (!recordSheet) {
      throw new Error(`'${RECORD_SHEET_NAME}' 탭을 찾을 수 없습니다.`);
    }
    
    const recordData = recordSheet.getDataRange().getValues();
    let records = [];
    
    // 1행(헤더)을 제외한 데이터가 있을 경우에만 읽기
    if (recordData.length > 1) {
      // 1행(헤더)에서 각 항목의 열 위치를 자동으로 탐색 (없으면 기본 위치 0, 1, 2, 3 사용)
      const recordHeaders = recordData[0].map(h => String(h).trim());
      const nameIdx = recordHeaders.indexOf('이름') !== -1 ? recordHeaders.indexOf('이름') : 0;
      const sportIdx = recordHeaders.indexOf('종목') !== -1 ? recordHeaders.indexOf('종목') : 1;
      const recordIdx = recordHeaders.indexOf('기록') !== -1 ? recordHeaders.indexOf('기록') : 2;
      const timeIdx = recordHeaders.indexOf('시간') !== -1 ? recordHeaders.indexOf('시간') : 3;

      const recordValues = recordData.slice(1);
      
      records = recordValues
        .filter(row => row[nameIdx] !== '' || row[sportIdx] !== '' || row[recordIdx] !== '') // 완전히 빈 행 제외
        .map((row, index) => {
          // 시간 데이터 포맷팅 (Date 객체인 경우 보기 좋은 문자열로 변환)
          let formattedTime = row[timeIdx];
          if (row[timeIdx] instanceof Date) {
            formattedTime = Utilities.formatDate(row[timeIdx], 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
          }
          
          return {
            rowId: index + 2, // 시트 상의 실제 행 번호
            name: String(row[nameIdx] || '').trim(),
            sport: String(row[sportIdx] || '').trim(),
            record: String(row[recordIdx] || '').trim(),
            time: String(formattedTime || '').trim()
          };
        })
        .reverse(); // 시트 아래쪽에 쌓인 최신 데이터를 배열 맨 앞(최신순)으로 역순 정렬
    }

    // 2. "학생명단" 시트 데이터 읽기 (행/열 자동 감지)
    const studentSheet = ss.getSheetByName(STUDENT_SHEET_NAME);
    if (!studentSheet) {
      throw new Error(`'${STUDENT_SHEET_NAME}' 탭을 찾을 수 없습니다.`);
    }
    
    // 시트 전체 사용 범위 데이터를 가져와 행/열 방향 및 헤더 위치를 자동 분석
    const studentData = studentSheet.getDataRange().getValues();
    let students = [];
    
    if (studentData.length > 0) {
      const firstRow = studentData[0].map(cell => String(cell).trim());
      const firstCol = studentData.map(row => String(row[0]).trim());
      
      // 1열(A열)에 '이름' 헤더가 있는 경우 -> 가로 방향(행/열 바뀜) 시트로 판단
      const isTransposed = !firstRow.includes('이름') && firstCol.includes('이름');
      
      if (isTransposed) {
        // 가로 방향 시트 처리 (예: 1행 번호, 2행 성별, 3행 이름)
        const numRowIdx = firstCol.indexOf('번호');
        const genderRowIdx = firstCol.indexOf('성별');
        const nameRowIdx = firstCol.indexOf('이름');
        
        const colCount = studentData[0].length;
        for (let c = 1; c < colCount; c++) {
          const nameVal = String(studentData[nameRowIdx][c] || '').trim();
          if (nameVal !== '') {
            students.push({
              number: numRowIdx !== -1 ? studentData[numRowIdx][c] : c,
              gender: genderRowIdx !== -1 ? String(studentData[genderRowIdx][c] || '').trim() : '',
              name: nameVal
            });
          }
        }
      } else {
        // 일반적인 세로 방향 시트 처리 (1행 헤더에서 '이름', '번호', '성별' 열 위치 자동 탐색)
        const numIdx = firstRow.indexOf('번호');
        const genderIdx = firstRow.indexOf('성별');
        let nameIdx = firstRow.indexOf('이름');
        
        // 만약 '이름' 헤더를 못 찾았다면, '남'/'여'가 아닌 한글 이름이 들어있는 열을 자동 추론
        if (nameIdx === -1 && studentData.length > 1) {
          const sampleRow = studentData[1];
          nameIdx = sampleRow.findIndex(val => {
            const str = String(val).trim();
            return isNaN(str) && str !== '남' && str !== '여' && str !== '남자' && str !== '여자' && str !== '';
          });
          if (nameIdx === -1) nameIdx = 1;
        }
        
        // 1행이 헤더인지 확인 (첫 행에 '이름'이나 '번호'가 있으면 2행부터 데이터로 간주)
        const startRow = (firstRow.includes('이름') || firstRow.includes('번호')) ? 1 : 0;
        
        for (let r = startRow; r < studentData.length; r++) {
          const row = studentData[r];
          const nameVal = String(row[nameIdx] || '').trim();
          if (nameVal !== '') {
            students.push({
              number: numIdx !== -1 ? row[numIdx] : r,
              gender: genderIdx !== -1 ? String(row[genderIdx] || '').trim() : '',
              name: nameVal
            });
          }
        }
      }
    }

    // 성공 응답 반환
    return createJsonResponse({
      status: 'success',
      students: students,
      records: records
    });

  } catch (error) {
    // 에러 발생 시에도 JSON 형태로 반환
    return createJsonResponse({
      status: 'error',
      message: error.message
    });
  }
}

/**
 * POST 요청 처리: "기록" 탭에 새로운 운동 기록 행을 추가합니다.
 * 
 * [CORS Preflight(OPTIONS) 회피 핵심 설명]
 * 브라우저에서 fetch로 POST 요청 시 'application/json' 헤더를 붙이면 OPTIONS 요청이 먼저 발생하여
 * Apps Script에서 CORS 에러가 납니다.
 * 따라서 프론트엔드에서는 headers 없이 body에 JSON.stringify({...})를 담아 보내면(text/plain 처리),
 * e.postData.contents로 안전하게 파싱할 수 있습니다.
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('전달된 데이터가 없습니다.');
    }

    // text/plain 또는 application/x-www-form-urlencoded 등 다양한 형태로 들어올 수 있는 데이터 파싱
    let payload;
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      // 만약 form-data 파라미터(e.parameter)로 전달된 경우 대비
      payload = e.parameter;
    }

    const name = payload.name ? String(payload.name).trim() : '';
    const sport = payload.sport ? String(payload.sport).trim() : '';
    const record = payload.record ? String(payload.record).trim() : '';

    if (!name || !sport || !record) {
      throw new Error('필수 항목(이름, 종목, 기록)이 누락되었습니다.');
    }

    // 서버 기준 현재 한국 시간(KST) 자동 생성
    const now = new Date();
    const serverTime = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const recordSheet = ss.getSheetByName(RECORD_SHEET_NAME);
    
    if (!recordSheet) {
      throw new Error(`'${RECORD_SHEET_NAME}' 탭을 찾을 수 없습니다.`);
    }

    // "기록" 시트의 1행(헤더) 순서를 확인해 열이 바뀌었더라도 맞는 위치에 저장
    const lastCol = Math.max(recordSheet.getLastColumn(), 4);
    const headers = recordSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
    
    const nameIdx = headers.indexOf('이름') !== -1 ? headers.indexOf('이름') : 0;
    const sportIdx = headers.indexOf('종목') !== -1 ? headers.indexOf('종목') : 1;
    const recordIdx = headers.indexOf('기록') !== -1 ? headers.indexOf('기록') : 2;
    const timeIdx = headers.indexOf('시간') !== -1 ? headers.indexOf('시간') : 3;

    const newRow = new Array(lastCol).fill('');
    newRow[nameIdx] = name;
    newRow[sportIdx] = sport;
    newRow[recordIdx] = record;
    newRow[timeIdx] = serverTime;

    // 시트 맨 아래에 새 행 추가
    recordSheet.appendRow(newRow);

    return createJsonResponse({
      status: 'success',
      message: '기록이 성공적으로 저장되었습니다.',
      data: {
        name: name,
        sport: sport,
        record: record,
        time: serverTime
      }
    });

  } catch (error) {
    return createJsonResponse({
      status: 'error',
      message: error.message
    });
  }
}

/**
 * Google Apps Script 표준 JSON 응답 헬퍼 함수
 * MimeType.JSON을 사용하면 Apps Script가 자동으로 적절한 Access-Control-Allow-Origin 헤더를 처리합니다.
 */
function createJsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}