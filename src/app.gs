// ===== 設定エリア =====
// 以下のIDは実際のファイルIDに置き換えてください
const FORM_ID = 'YOUR_GOOGLE_FORM_ID_HERE';
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
const PRESENTATION_ID = 'YOUR_PRESENTATION_ID_HERE'; // 既存のスライドIDがあれば設定、なければ新規作成

// ===== 投票アンケート自動更新 =====
function updateForm() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Master');
    
    if (!sheet) {
      throw new Error('Masterシートが見つかりません。');
    }
    
    // 登壇チーム情報を取得（A列・B列）
    const teamsRange = sheet.getRange('A2:B').getValues();
    const teams = teamsRange.filter(row => row[0] !== '' && row[1] !== '');
    
    // 評価項目を取得（C列）
    const criteriaRange = sheet.getRange('C2:C').getValues();
    const criteria = criteriaRange.filter(row => row[0] !== '').map(row => row[0]);
    
    // チーム名のリストから重複を除去
    const teamNames = [...new Set(teams.map(team => team[0]))];
    
    if (teams.length === 0 || criteria.length === 0) {
      throw new Error('Masterシートのデータが不足しています。');
    }
    
    // フォームを取得・更新
    const form = FormApp.openById(FORM_ID);
    form.getItems().forEach(item => form.deleteItem(item));
    
    form.setTitle('イベントアンケート - 登壇者評価');
    form.setDescription('各登壇者の発表について評価をお願いします。');
    
    // 参加者氏名入力欄
    form.addTextItem()
        .setTitle('参加者氏名 or ハンドルネーム')
        .setRequired(true);
    
    // チーム選択（ラジオボタン）
    const teamRadio = form.addMultipleChoiceItem()
        .setTitle('所属チーム')
        .setRequired(true);
    
    const teamChoices = teamNames.map(team => teamRadio.createChoice(team));
    if (!teamNames.includes('その他')) {
      teamChoices.push(teamRadio.createChoice('その他'));
    }
    teamRadio.setChoices(teamChoices);
    
    // 各登壇チームのグリッド評価
    teams.forEach(([teamName, title]) => {
      form.addGridItem()
          .setTitle(teamName)
          .setHelpText(title)
          .setRows(criteria)
          .setColumns(['1 (普通)', '2', '3 (良い)'])
          .setRequired(true);
    });
    
    console.log('✅ フォームの更新が完了しました');
    
  } catch (error) {
    console.error('フォーム更新エラー:', error.message);
  }
}

// ===== アンケート結果集計関数 =====
function aggregateSurveyResults() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const responseSheet = ss.getSheetByName('Result');
    
    if (!responseSheet || responseSheet.getLastRow() <= 1) {
      throw new Error('回答データがありません');
    }
    
    const headers = responseSheet.getRange(1, 1, 1, responseSheet.getLastColumn()).getValues()[0];
    const data = responseSheet.getRange(2, 1, responseSheet.getLastRow() - 1, responseSheet.getLastColumn()).getValues();
    
    // ヘッダーからチーム評価列を特定
    const teamColumns = [];
    const nameColumnIndex = headers.findIndex(h => h.includes('名前') || h.includes('ハンドルネーム'));
    const teamColumnIndex = headers.findIndex(h => h.includes('所属チーム'));
    
    headers.forEach((header, index) => {
      if (index > 2) {
        const match = header.match(/^(.+?)\s*\[(.+?)\]$/);
        if (match) {
          const teamName = match[1].trim();
          const criterion = match[2].trim();
          
          let teamData = teamColumns.find(tc => tc.teamName === teamName);
          if (!teamData) {
            teamData = { teamName: teamName, columns: [] };
            teamColumns.push(teamData);
          }
          teamData.columns.push({ columnIndex: index, criterion: criterion });
        }
      }
    });
    
    // 集計処理
    const teamResults = {};
    teamColumns.forEach(teamColumn => {
      const teamName = teamColumn.teamName;
      teamResults[teamName] = {
        totalPoints: 0,
        totalPointsExcludingSelf: 0,
        responseCount: 0,
        responseCountExcludingSelf: 0
      };
    });
    
    data.forEach((row) => {
      const respondentTeam = row[teamColumnIndex] || '';
      
      teamColumns.forEach(teamColumn => {
        const teamName = teamColumn.teamName;
        const isOwnTeam = respondentTeam.includes(teamName) || teamName.includes(respondentTeam);
        
        let teamTotalPoints = 0;
        teamColumn.columns.forEach(col => {
          const cellValue = row[col.columnIndex];
          let points = 0;
          
          if (typeof cellValue === 'string') {
            if (cellValue.includes('1') || cellValue.includes('普通')) points = 1;
            else if (cellValue.includes('2')) points = 2;
            else if (cellValue.includes('3') || cellValue.includes('良い')) points = 3;
          } else if (typeof cellValue === 'number') {
            points = Math.min(Math.max(cellValue, 1), 3);
          }
          
          teamTotalPoints += points;
        });
        
        teamResults[teamName].totalPoints += teamTotalPoints;
        teamResults[teamName].responseCount++;
        
        if (!isOwnTeam) {
          teamResults[teamName].totalPointsExcludingSelf += teamTotalPoints;
          teamResults[teamName].responseCountExcludingSelf++;
        }
      });
    });
    
    // 結果をシートに出力
    const existingSheet = ss.getSheetByName('集計結果');
    if (existingSheet) ss.deleteSheet(existingSheet);
    
    const resultSheet = ss.insertSheet('集計結果');
    const headers2 = ['チーム名', '全体合計ポイント', '自チーム除外合計ポイント', '回答者数（全体）', '回答者数（自チーム除外）'];
    resultSheet.getRange(1, 1, 1, headers2.length).setValues([headers2]);
    resultSheet.getRange(1, 1, 1, headers2.length).setFontWeight('bold').setBackground('#f0f0f0');
    
    const outputData = Object.entries(teamResults).map(([teamName, result]) => [
      teamName,
      result.totalPoints,
      result.totalPointsExcludingSelf,
      result.responseCount,
      result.responseCountExcludingSelf
    ]);
    
    if (outputData.length > 0) {
      resultSheet.getRange(2, 1, outputData.length, headers2.length).setValues(outputData);
    }
    
    console.log('✅ 集計結果を出力しました');
    
  } catch (error) {
    console.error('集計エラー:', error.message);
  }
}

// ===== 結果発表スライド作成関数 =====
function createResultPresentation() {
  try {
    // 集計実行
    aggregateSurveyResults();
    
    // 集計データ取得
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const resultSheet = ss.getSheetByName('集計結果');
    const masterSheet = ss.getSheetByName('Master');
    
    if (!resultSheet) throw new Error('集計結果シートが見つかりません');
    
    const data = resultSheet.getRange(2, 1, resultSheet.getLastRow() - 1, 5).getValues();
    
    // 登壇タイトル取得
    const teamTitles = {};
    if (masterSheet) {
      const masterData = masterSheet.getRange('A2:B').getValues();
      masterData.forEach(row => {
        if (row[0] && row[1]) teamTitles[row[0]] = row[1];
      });
    }
    
    // ランキング作成（全体ポイントでソート）
    console.log('ソート前のデータ:', data.map(row => `${row[0]}: 全体${row[1]}点, 自チーム除外${row[2]}点`));
    
    const ranking = data.map(row => ({
      teamName: row[0],
      totalPoints: Number(row[1]), // 数値に変換
      pointsExcludingSelf: Number(row[2]), // 数値に変換
      responseCount: row[3],
      responseCountExcludingSelf: row[4],
      title: teamTitles[row[0]] || 'タイトル未設定'
    })).sort((a, b) => b.totalPoints - a.totalPoints)
      .map((team, index) => ({ rank: index + 1, ...team }));
    
    console.log('ソート後のランキング:', ranking.map(team => `${team.rank}位 ${team.teamName}: ${team.totalPoints}点`));
    
    // プレゼンテーション作成
    let presentation;
    if (PRESENTATION_ID) {
      presentation = SlidesApp.openById(PRESENTATION_ID);
      // 既存スライドを削除
      const slides = presentation.getSlides();
      for (let i = slides.length - 1; i > 0; i--) {
        slides[i].remove();
      }
      if (slides.length > 0) {
        slides[0].getPageElements().forEach(element => element.remove());
      }
    } else {
      presentation = SlidesApp.create('結果発表');
    }
    
    // タイトルスライド
    const slides = presentation.getSlides();
    let titleSlide = slides.length === 0 ? presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE) : slides[0];
    
    titleSlide.insertTextBox('結果発表', 50, 150, 600, 100)
              .getText().getTextStyle().setFontSize(48).setBold(true).setForegroundColor('#1a73e8');
    titleSlide.insertTextBox(`全${ranking.length}チームの評価結果`, 50, 280, 600, 50)
              .getText().getTextStyle().setFontSize(24).setForegroundColor('#5f6368');
    
    // 各順位のスライド（3位→2位→1位の順番で表示）
    const displayOrder = ranking.slice().reverse(); // 順位を逆順にする
    
    displayOrder.forEach((team) => {
      // 順位発表スライド
      const rankSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      let rankText, rankColor;
      
      if (team.rank === 1) {
        rankText = '🥇 優勝！';
        rankColor = '#ffd700';
      } else if (team.rank === 2) {
        rankText = '🥈 第2位！';
        rankColor = '#c0c0c0';
      } else if (team.rank === 3) {
        rankText = '🥉 第3位！';
        rankColor = '#cd7f32';
      } else {
        rankText = `第${team.rank}位`;
        rankColor = '#1a73e8';
      }
      
      rankSlide.insertTextBox(rankText, 50, 100, 600, 150)
               .getText().getTextStyle().setFontSize(72).setBold(true).setForegroundColor(rankColor);
      rankSlide.insertTextBox(`${team.totalPoints}ポイント`, 50, 280, 600, 80)
               .getText().getTextStyle().setFontSize(36).setForegroundColor('#5f6368');
      
      // チーム詳細スライド
      const detailSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      
      // チーム名（タイトル）+ 獲得ポイント
      const teamTitleText = `${team.teamName}\n${team.totalPoints}ポイント`;
      detailSlide.insertTextBox(teamTitleText, 50, 100, 600, 120)
                 .getText().getTextStyle().setFontSize(44).setBold(true).setForegroundColor('#1a73e8');
      
      // 登壇タイトル
      detailSlide.insertTextBox(team.title, 50, 250, 600, 100)
                 .getText().getTextStyle().setFontSize(32).setForegroundColor('#202124');
    });
    
    // 全チーム得票数表示スライド（表形式）
    const allResultsSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    
    // タイトル
    allResultsSlide.insertTextBox('全チーム得票数', 50, 30, 600, 60)
                   .getText().getTextStyle().setFontSize(36).setBold(true).setForegroundColor('#1a73e8');
    
    // 表を作成
    const numRows = ranking.length + 1; // データ行 + ヘッダー行
    const numCols = 3; // 順位、チーム名、ポイント
    
    const table = allResultsSlide.insertTable(numRows, numCols, 50, 120, 600, numRows * 40);
    
    // ヘッダー行を設定
    const headerRow = table.getRow(0);
    headerRow.getCell(0).getText().setText('順位').getTextStyle().setBold(true).setFontSize(18);
    headerRow.getCell(1).getText().setText('チーム名').getTextStyle().setBold(true).setFontSize(18);
    headerRow.getCell(2).getText().setText('ポイント').getTextStyle().setBold(true).setFontSize(18);
    
    // ヘッダー行の背景色を設定
    for (let col = 0; col < numCols; col++) {
      headerRow.getCell(col).getFill().setSolidFill('#e8f0fe');
    }
    
    // データ行を設定
    ranking.forEach((team, index) => {
      const row = table.getRow(index + 1);
      
      // 順位（メダル付き）
      const medal = team.rank <= 3 ? ['🥇', '🥈', '🥉'][team.rank - 1] : '';
      const rankText = `${medal} ${team.rank}位`;
      row.getCell(0).getText().setText(rankText).getTextStyle().setFontSize(16);
      
      // チーム名
      row.getCell(1).getText().setText(team.teamName).getTextStyle().setFontSize(16);
      
      // ポイント
      row.getCell(2).getText().setText(`${team.totalPoints}点`).getTextStyle().setFontSize(16);
      
      // 上位3位の行に色付け
      if (team.rank <= 3) {
        const colors = ['#fff2cc', '#f4f4f4', '#fce5cd']; // ゴールド、シルバー、ブロンズ色
        for (let col = 0; col < numCols; col++) {
          row.getCell(col).getFill().setSolidFill(colors[team.rank - 1]);
        }
      }
    });
    
    // サマリースライド
    const summarySlide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    
    // タイトル
    summarySlide.insertTextBox('最終結果', 50, 50, 600, 80)
                .getText().getTextStyle().setFontSize(44).setBold(true).setForegroundColor('#1a73e8');
    
    // 終了スライド（感謝メッセージ）
    const endingSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    
    // タイトル
    endingSlide.insertTextBox('結果発表 終了', 50, 150, 600, 80)
               .getText().getTextStyle().setFontSize(48).setBold(true).setForegroundColor('#1a73e8');
    
    // 感謝メッセージ
    endingSlide.insertTextBox('🎉 みなさま、お疲れさまでした！', 50, 250, 600, 100)
               .getText().getTextStyle().setFontSize(32).setForegroundColor('#202124');
    
    console.log('✅ 結果発表スライドを作成しました');
    console.log('URL:', presentation.getUrl());
    
    return presentation;
    
  } catch (error) {
    console.error('スライド作成エラー:', error.message);
  }
}