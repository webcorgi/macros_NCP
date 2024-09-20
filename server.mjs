import XLSX from 'xlsx';
import keySender from 'node-key-sender';
import clipboardy from 'clipboardy';

/**
 * @name 키보드매크로
 * @description 유저 인증차단해제 (NCP용).
 * 단순 키보드 조작만 가능하며 html 정보 접근은 할 수 없다.
 * @author 김대현
 * @date 24-09-20
 */
const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

const pressKey = async (key) => {
    return keySender.sendKey(key);
};

const runMacro = async () => {
    // Excel 파일 읽기
    const workbook = XLSX.readFile('excel.xlsx');
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    let rowIndex = 1;  // 엑셀의 행 번호 (1부터 시작)

     // 2. Alt+Tab을 눌러 브라우저 창으로 이동
    await keySender.sendCombination(['alt', 'tab']);
     await sleep(50);  // 창 전환 대기

    while (true) {
        const cellAddress = 'A' + rowIndex;
        const cell = sheet[cellAddress];

        if (!cell || !cell.v) {
            console.log("빈 셀 발견 (전체 작업 완료). 매크로를 종료합니다.");
            break;
        }

        const value = cell.v.toString();
        console.log(`처리 중인 이메일 (A${rowIndex}): ${value}`);

        // 1. 엑셀 셀 내용을 클립보드에 복사 (A1부터 차례대로)
        await clipboardy.write(value);
        await sleep(1); // 올바른 동작처리를 위한 공백

        // 2. 복사된 정보 붙여넣기
        await keySender.sendCombination(['control', 'v']);
        await sleep(1);

        // 3. 이메일 정보를 검색
        await pressKey('enter');
        await sleep(1);

        // 4. 차단해제버튼으로 이동
        for (let i = 0; i < 4; i++) {
            await pressKey('tab');
            await sleep(1);
        }

        // 5. 버튼 클릭 (차단해제 아니라면 조회버튼 클릭해서 확인하고 넘어감)
        await pressKey('enter');
        await sleep(1);

        // 6. 열린 팝업창 닫기
        for (let i = 0; i < 2; i++) {
            await pressKey('tab');
            await sleep(1);
        }
        await pressKey('enter');
        await sleep(1);

        // 7. 다시 브라우저의 이메일 입력 필드로 돌아가기
        for (let i = 0; i < 8; i++) {
            await pressKey('tab');
            await sleep(1);
        }
        rowIndex++;  // 다음 행으로
    }

    console.log("매크로 실행 완료");
};

runMacro().catch(console.error);