/** 中史 UT1 v2：Google Sheets 排行榜。先執行 setupLeaderboard，再部署網頁應用程式。 */
const BOARD_VERSION = 'ut1-v2-20261027';
const RUN_HEADERS = ['run_id','started_at','finished_at','version','grade','player_hash','alias','deck_ids','status','correct','seconds'];
const BEST_HEADERS = ['version','grade','player_hash','alias','correct','seconds','finished_at','run_id'];

function doGet(e) {
  if (e && e.parameter && e.parameter.view === 'bridge') {
    const origin = String(e.parameter.origin || '');
    const nonce = String(e.parameter.nonce || '');
    if (origin !== 'https://vincentttchan.github.io' || !/^[a-zA-Z0-9-]{20,80}$/.test(nonce))
      return HtmlService.createHtmlOutput('排行榜連線來源無效。');
    const template = HtmlService.createTemplateFromFile('Bridge');
    template.bootJson = JSON.stringify({origin,nonce,channel:'ut1-leaderboard-bridge-v1'});
    return template.evaluate().setTitle('排行榜連線').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('index').setTitle('中史溫習百萬富翁｜中一・中二');
}

// 只需由擁有者在編輯器執行一次；不會公開試算表或刪改其他檔案。
function setupLeaderboard() {
  const lock = LockService.getScriptLock();
  lock.waitLock(8000);
  try {
    const props = PropertiesService.getScriptProperties();
    let id = props.getProperty('MONOPOLY_SHEET_ID');
    if (!id) {
      const ss = SpreadsheetApp.create('中史UT1百萬富翁－排行榜');
      const runs = ss.getSheets()[0]; runs.setName('Runs');
      runs.getRange(1,1,1,RUN_HEADERS.length).setValues([RUN_HEADERS]);
      ss.insertSheet('BestScores').getRange(1,1,1,BEST_HEADERS.length).setValues([BEST_HEADERS]);
      id = ss.getId();
      props.setProperty('MONOPOLY_SHEET_ID', id);
      props.setProperty('MONOPOLY_SALT', Utilities.getUuid());
    }
    const ss = SpreadsheetApp.openById(id);
    ['Runs','BestScores'].forEach(name => {
      const sheet = ss.getSheetByName(name);
      sheet.setFrozenRows(1);
      sheet.getRange(1,1,1,sheet.getLastColumn()).setBackground('#172554').setFontColor('#fde047').setFontWeight('bold');
      sheet.autoResizeColumns(1,sheet.getLastColumn());
    });
    console.log('私人試算表：' + ss.getUrl());
    return '排行榜設定完成。請查看執行記錄中的試算表連結。';
  } finally { lock.releaseLock(); }
}

function store_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('MONOPOLY_SHEET_ID');
  if (!id || !p.getProperty('MONOPOLY_SALT')) throw new Error('老師尚未執行 setupLeaderboard。');
  return SpreadsheetApp.openById(id);
}
function validateGrade_(grade) {
  if (grade !== 's1' && grade !== 's2') throw new Error('請選擇中一或中二。');
}
function alias_(value) {
  const alias = String(value || '').normalize('NFKC').trim();
  if (!/^[\p{L}\p{N}][\p{L}\p{N} _·-]{1,15}$/u.test(alias))
    throw new Error('暱稱須為2–16個中文字、英文字母或數字，可含空格、短橫線；請勿填寫真名。');
  return alias;
}
function hashPlayer_(token) {
  if (typeof token !== 'string' || !/^[A-Za-z0-9-]{20,80}$/.test(token)) throw new Error('玩家識別碼無效，請重新整理頁面。');
  const salt = PropertiesService.getScriptProperties().getProperty('MONOPOLY_SALT');
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + token, Utilities.Charset.UTF_8)
    .map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}
function shuffle_(items) {
  const a = items.slice();
  for (let i=a.length-1;i>0;i--) { const j=Math.floor(Math.random()*(i+1)); const t=a[i];a[i]=a[j];a[j]=t; }
  return a;
}
function rows_(sheet) {
  const count = sheet.getLastRow();
  return count < 2 ? [] : sheet.getRange(2,1,count-1,sheet.getLastColumn()).getValues();
}
function startRankedRun(input) {
  input = input || {};
  validateGrade_(input.grade);
  if (input.version !== BOARD_VERSION) throw new Error('題庫版本已更新，請重新整理網頁。');
  const ss = store_(), nickname = alias_(input.alias), player = hashPlayer_(input.playerToken);
  const lock=LockService.getScriptLock(); lock.waitLock(8000);
  try {
    const cache=CacheService.getScriptCache(), throttle='start:' + player;
    const starts=Number(cache.get(throttle)||0);
    if(starts >= 3) throw new Error('開始次數太密，請等一分鐘再試。');
    cache.put(throttle,String(starts+1),60);
    const bank = getQuestionBanks_()[input.grade];
    const deck = ['easy','medium','hard'].reduce((all,diff) => all.concat(shuffle_(bank[diff]).slice(0,10)),[]);
    const runId=Utilities.getUuid(), now=Date.now();
    ss.getSheetByName('Runs').appendRow([runId,now,'',BOARD_VERSION,input.grade,player,nickname,JSON.stringify(deck.map(q=>q.id)),'pending','','']);
    SpreadsheetApp.flush();
    return {runId,version:BOARD_VERSION,deck:deck.map(q => ({id:q.id,options:shuffle_(q.options)}))};
  } finally { lock.releaseLock(); }
}
function compareScore_(a,b) { return Number(b.correct)-Number(a.correct) || Number(a.seconds)-Number(b.seconds); }
function leaderboard_(ss,grade) {
  const entries=rows_(ss.getSheetByName('BestScores'))
    .filter(r => r[0]===BOARD_VERSION && r[1]===grade)
    .map(r => ({alias:String(r[3]),correct:Number(r[4]),seconds:Number(r[5])}))
    .sort((a,b)=>compareScore_(a,b)||a.alias.localeCompare(b.alias));
  let previous,rank=0;
  entries.forEach((e,i)=>{ if(!previous || compareScore_(previous,e)!==0) rank=i+1; e.rank=rank;previous=e; });
  return {version:BOARD_VERSION,grade,entries:entries.slice(0,20),totalPlayers:entries.length,updatedAt:Date.now()};
}
function getLeaderboard(grade) { validateGrade_(grade);return leaderboard_(store_(),grade); }
function submitRankedRun(input) {
  input=input||{};
  if(typeof input.runId!=='string'||input.runId.length>80)throw new Error('局次識別碼無效。');
  const ss=store_(),player=hashPlayer_(input.playerToken);
  const lock=LockService.getScriptLock();lock.waitLock(8000);
  try {
    const sheet=ss.getSheetByName('Runs'),rows=rows_(sheet),idx=rows.findIndex(r=>r[0]===input.runId&&r[5]===player);
    if(idx<0)throw new Error('找不到這一局，請重新開始。');
    const run=rows[idx];
    if(run[3]!==BOARD_VERSION)throw new Error('本局題庫版本已過期。');
    if(run[8]==='complete')return {correct:Number(run[9]),seconds:Number(run[10]),board:leaderboard_(ss,run[4]),duplicate:true};
    const elapsed=Date.now()-Number(run[1]);
    if(elapsed<0||elapsed>4*60*60*1000)throw new Error('本局已超過4小時，請重新開始。');
    const ids=JSON.parse(run[7]),answers=input.answers;
    if(!Array.isArray(answers)||answers.length<1||answers.length>30)throw new Error('作答紀錄不完整。');
    const bank=getQuestionBanks_()[run[4]],byId={};
    ['easy','medium','hard'].forEach(d=>bank[d].forEach(q=>byId[q.id]=q));
    let correct=0,wrong=false;
    answers.forEach((a,i)=>{
      if(!a||a.id!==ids[i]||typeof a.answer!=='string'||!byId[a.id].options.includes(a.answer)||wrong)
        throw new Error('作答次序或內容無效。');
      if(a.answer===byId[a.id].options[byId[a.id].ans])correct++;else wrong=true;
    });
    if(!wrong&&correct!==30)throw new Error('必須完成本局才能登榜。');
    const seconds=Math.max(1,Math.ceil(elapsed/1000)),now=Date.now();
    const best=ss.getSheetByName('BestScores'),bestRows=rows_(best);
    const bidx=bestRows.findIndex(r=>r[0]===BOARD_VERSION&&r[1]===run[4]&&r[2]===player);
    const candidate={correct,seconds};
    const improved=bidx<0||compareScore_(candidate,{correct:bestRows[bidx][4],seconds:bestRows[bidx][5]})<0;
    if(improved){
      const record=[BOARD_VERSION,run[4],player,run[6],correct,seconds,now,input.runId];
      if(bidx<0)best.appendRow(record);else best.getRange(bidx+2,1,1,BEST_HEADERS.length).setValues([record]);
    }
    // Best score is written first. A failed write can be retried idempotently using the same run ID.
    sheet.getRange(idx+2,3).setValue(now);
    sheet.getRange(idx+2,9,1,3).setValues([['complete',correct,seconds]]);
    SpreadsheetApp.flush();
    return {correct,seconds,improved,board:leaderboard_(ss,run[4]),duplicate:false};
  } finally { lock.releaseLock(); }
}

/** 修訂題庫 v2，與網頁版本一致。 */
function getQuestionBanks_() { return {
  "s1": {
    "easy": [
      {
        "id": "s1-easy-01",
        "q": "依據筆記，舊石器時代先民主要居於哪裏？",
        "options": [
          "洞穴",
          "木造房屋",
          "夯土宮室",
          "磚造民居"
        ],
        "ans": 0,
        "explanation": "筆記把洞穴居住列為舊石器時代的生活特徵。",
        "topic": "華夏起源"
      },
      {
        "id": "s1-easy-02",
        "q": "舊石器時代的石器主要有甚麼特點？",
        "options": [
          "粗製石器",
          "磨製石器",
          "青銅工具",
          "鐵製工具"
        ],
        "ans": 0,
        "explanation": "舊石器較粗製；新石器時代才發展經磨製的石器。",
        "topic": "華夏起源"
      },
      {
        "id": "s1-easy-03",
        "q": "哪項是筆記所列的新石器時代生產活動？",
        "options": [
          "耕作及畜牧",
          "採集及狩獵為主",
          "製作青銅禮器",
          "以鐵器耕田"
        ],
        "ans": 0,
        "explanation": "新石器先民懂耕作及畜牧，生活方式較前進步。",
        "topic": "華夏起源"
      },
      {
        "id": "s1-easy-04",
        "q": "相同血緣的先民聚居形成甚麼群體？",
        "options": [
          "氏族",
          "部落聯盟",
          "諸侯國",
          "王朝"
        ],
        "ans": 0,
        "explanation": "筆記區分氏族、部落及部落聯盟；氏族以相同血緣為基礎。",
        "topic": "華夏起源"
      },
      {
        "id": "s1-easy-05",
        "q": "「炎黃子孫」的「炎黃」指哪兩位始祖？",
        "options": [
          "炎帝和黃帝",
          "黃帝和堯",
          "堯和舜",
          "舜和禹"
        ],
        "ans": 0,
        "explanation": "炎帝、黃帝被尊為中華民族始祖。",
        "topic": "黃帝功績"
      },
      {
        "id": "s1-easy-06",
        "q": "傳說中，嫘祖的主要貢獻是甚麼？",
        "options": [
          "養蠶取絲",
          "創造文字",
          "創制音律",
          "製造指南車"
        ],
        "ans": 0,
        "explanation": "筆記把養蠶取絲歸於嫘祖；這屬傳說中的發明。",
        "topic": "黃帝功績"
      },
      {
        "id": "s1-easy-07",
        "q": "傳說中，黃帝哪位部下創造文字？",
        "options": [
          "倉頡",
          "嫘祖",
          "堯",
          "禹"
        ],
        "ans": 0,
        "explanation": "筆記記述倉頡造字；題幹保留「傳說中」以區別考證史實。",
        "topic": "黃帝功績"
      },
      {
        "id": "s1-easy-08",
        "q": "堯選擇哪位以孝聞名的人繼任？",
        "options": [
          "舜",
          "禹",
          "啟",
          "湯"
        ],
        "ans": 0,
        "explanation": "堯以舜為繼任人；舜後來再把共主之位交予禹。",
        "topic": "堯舜禹"
      },
      {
        "id": "s1-easy-09",
        "q": "禹主要因哪項功績獲舜選為共主？",
        "options": [
          "治水",
          "創造文字",
          "發明音律",
          "推翻暴君"
        ],
        "ans": 0,
        "explanation": "禹治水有功，是筆記所述獲舜選為共主的原因。",
        "topic": "堯舜禹"
      },
      {
        "id": "s1-easy-10",
        "q": "禹的兒子是誰？",
        "options": [
          "啟",
          "桀",
          "湯",
          "武庚"
        ],
        "ans": 0,
        "explanation": "啟繼位後，君位傳承由禪讓轉向世襲。",
        "topic": "夏朝"
      },
      {
        "id": "s1-easy-11",
        "q": "夏朝最後一任君主是誰？",
        "options": [
          "桀",
          "啟",
          "湯",
          "紂王"
        ],
        "ans": 0,
        "explanation": "桀是夏朝末代君主，後被湯推翻。",
        "topic": "夏朝"
      },
      {
        "id": "s1-easy-12",
        "q": "誰推翻夏朝，建立商朝？",
        "options": [
          "湯",
          "禹",
          "啟",
          "周武王"
        ],
        "ans": 0,
        "explanation": "湯推翻夏桀；周武王則是其後伐商。",
        "topic": "商朝"
      },
      {
        "id": "s1-easy-13",
        "q": "殷墟出土、可作商代文字證據的是甚麼？",
        "options": [
          "甲骨文",
          "小篆",
          "隸書",
          "楷書"
        ],
        "ans": 0,
        "explanation": "殷墟甲骨文是研究商代的重要文字材料。",
        "topic": "商朝"
      },
      {
        "id": "s1-easy-14",
        "q": "商朝最後一任君主是誰？",
        "options": [
          "紂王",
          "桀",
          "周武王",
          "周幽王"
        ],
        "ans": 0,
        "explanation": "夏末是桀、商末是紂王，兩者不可混淆。",
        "topic": "商朝"
      },
      {
        "id": "s1-easy-15",
        "q": "哪組國家全屬戰國七雄？",
        "options": [
          "韓、趙、魏",
          "韓、趙、晉",
          "齊、秦、周",
          "楚、燕、宋"
        ],
        "ans": 0,
        "explanation": "七雄為韓、趙、魏、楚、燕、齊、秦。",
        "topic": "春秋戰國"
      }
    ],
    "medium": [
      {
        "id": "s1-medium-01",
        "q": "哪項正確比較兩個石器時代的工具？",
        "options": [
          "舊石器較粗製，新石器多經磨製",
          "舊石器多經磨製，新石器較粗製",
          "舊石器以銅器為主，新石器以鐵器為主",
          "舊石器以鐵器為主，新石器以銅器為主"
        ],
        "ans": 0,
        "explanation": "筆記以石器製作技術區分兩個時代；銅器、鐵器並非此題的比較重點。",
        "topic": "華夏起源"
      },
      {
        "id": "s1-medium-02",
        "q": "先民群居如何有助提高生產力？",
        "options": [
          "透過分工合作完成生產活動",
          "由每個家庭各自負責全部工作",
          "減少成員之間的合作與交流",
          "以血緣遠近代替工作分配"
        ],
        "ans": 0,
        "explanation": "群居兼具安全與分工合作的好處。",
        "topic": "華夏起源"
      },
      {
        "id": "s1-medium-03",
        "q": "部落聯盟推選的共同領袖稱為甚麼？",
        "options": [
          "共主",
          "諸侯",
          "天子使臣",
          "氏族成員"
        ],
        "ans": 0,
        "explanation": "氏族組成部落，部落結成聯盟並推選共主。",
        "topic": "華夏起源"
      },
      {
        "id": "s1-medium-04",
        "q": "傳說中黃帝被尊為「人文初祖」，與哪項貢獻最相關？",
        "options": [
          "多項發明改善先民生活",
          "把君位傳給自己的兒子",
          "以疏導方法治理洪水",
          "平定周代三監之亂"
        ],
        "ans": 0,
        "explanation": "筆記把黃帝的發明與改善先民生活聯繫起來。",
        "topic": "黃帝功績"
      },
      {
        "id": "s1-medium-05",
        "q": "把領袖位置傳給具才德的人，稱為甚麼制度？",
        "options": [
          "禪讓制",
          "世襲制",
          "分封制",
          "攝政制"
        ],
        "ans": 0,
        "explanation": "禪讓重才德；世襲著重家族內傳承。",
        "topic": "堯舜禹"
      },
      {
        "id": "s1-medium-06",
        "q": "哪項配對符合筆記對領袖傳承的描述？",
        "options": [
          "禪讓制—公天下",
          "禪讓制—家天下",
          "世襲制—公天下",
          "攝政制—禪讓天下"
        ],
        "ans": 0,
        "explanation": "筆記以「公天下」概括禪讓，以「家天下」概括世襲。",
        "topic": "堯舜禹"
      },
      {
        "id": "s1-medium-07",
        "q": "禹以疏導法治水，主要做法是甚麼？",
        "options": [
          "開通河道，把洪水引向大海",
          "加高堤壩，把洪水堵在原地",
          "遷走居民，放棄整理河道",
          "封閉河口，減少河水外流"
        ],
        "ans": 0,
        "explanation": "疏導著重開通河道，與堵塞的治水思路不同。",
        "topic": "夏朝"
      },
      {
        "id": "s1-medium-08",
        "q": "啟繼位，反映領袖傳承方式出現甚麼轉變？",
        "options": [
          "由禪讓轉為世襲",
          "由世襲轉為禪讓",
          "由分封轉為攝政",
          "由攝政轉為分封"
        ],
        "ans": 0,
        "explanation": "啟是禹之子；父傳子標誌「家天下」的形成。",
        "topic": "夏朝"
      },
      {
        "id": "s1-medium-09",
        "q": "依據筆記，商代被視為「信史」開始，理由是甚麼？",
        "options": [
          "有甲骨文提供文字證據",
          "傳說記載了黃帝多項發明",
          "君主曾多次遷移首都",
          "商朝已有軍隊及刑法"
        ],
        "ans": 0,
        "explanation": "此處判準是文字證據；國家制度或傳說本身不是同一種證據。",
        "topic": "商朝"
      },
      {
        "id": "s1-medium-10",
        "q": "筆記所述商朝多次遷都的原因之一是甚麼？",
        "options": [
          "逃避黃河水患",
          "減弱諸侯對周室的威脅",
          "監視殷遺民的反叛活動",
          "避開犬戎攻入鎬京的路線"
        ],
        "ans": 0,
        "explanation": "筆記把商代遷都與黃河水患相聯繫，最後定都殷。",
        "topic": "商朝"
      },
      {
        "id": "s1-medium-11",
        "q": "周武王封武庚於殷，主要希望達到甚麼目的？",
        "options": [
          "安撫殷遺民，減少敵意",
          "讓武庚代替成王攝政",
          "令殷遺民遷往洛邑",
          "讓三監歸順商朝王室"
        ],
        "ans": 0,
        "explanation": "保留商王後裔的地位，有助安撫殷遺民。",
        "topic": "周代分封"
      },
      {
        "id": "s1-medium-12",
        "q": "周代「三監」是指哪三人？",
        "options": [
          "管、蔡、霍",
          "管、蔡、武庚",
          "周公、成王、武庚",
          "申侯、幽王、平王"
        ],
        "ans": 0,
        "explanation": "筆記列管、蔡、霍為三監，職責是監視殷地。",
        "topic": "周代分封"
      },
      {
        "id": "s1-medium-13",
        "q": "周武王死後，誰因成王年幼而攝政？",
        "options": [
          "周公",
          "武庚",
          "管叔",
          "申侯"
        ],
        "ans": 0,
        "explanation": "成王年幼，由周公攝政；三監因此不滿。",
        "topic": "周代分封"
      },
      {
        "id": "s1-medium-14",
        "q": "周平王東遷到哪個都城？",
        "options": [
          "洛邑",
          "鎬京",
          "殷",
          "汴京"
        ],
        "ans": 0,
        "explanation": "平王由殘破的鎬京東遷洛邑，開始東周時期。",
        "topic": "西周與東周"
      },
      {
        "id": "s1-medium-15",
        "q": "齊桓公以哪個口號號令諸侯？",
        "options": [
          "尊王攘夷",
          "強榦弱枝",
          "重文輕武",
          "富國強兵"
        ],
        "ans": 0,
        "explanation": "「尊王」是尊重周室，「攘夷」是驅除外族入侵。",
        "topic": "春秋戰國"
      }
    ],
    "hard": [
      {
        "id": "s1-hard-01",
        "q": "【情境題】不同地域的先民保留各自風俗，又共同形成中華文明。哪項概括最恰當？",
        "options": [
          "多元文化共同形成整體",
          "各地文化發展互不相關",
          "各地風俗由一個氏族統一定下",
          "中華文明只來自同一種生活方式"
        ],
        "ans": 0,
        "explanation": "多元一體同時強調地域差異與共同形成的文明整體。",
        "topic": "華夏起源"
      },
      {
        "id": "s1-hard-02",
        "q": "【資料判讀】甲：居洞穴、採集狩獵；乙：建房屋、耕作畜牧。哪項推論最有根據？",
        "options": [
          "乙較甲發展出定居生產的生活",
          "甲較乙掌握更成熟的農業技術",
          "乙較甲更依賴野生動植物採集",
          "甲乙的生產方式並無明顯分別"
        ],
        "ans": 0,
        "explanation": "耕作畜牧及建造房屋，支持生活方式轉向較穩定的定居生產。",
        "topic": "華夏起源"
      },
      {
        "id": "s1-hard-03",
        "q": "【比較題】禪讓制與世襲制的主要分別，應從哪方面比較？",
        "options": [
          "繼任者的選取依據",
          "君主居住的宮室規模",
          "王朝所處的河流位置",
          "軍隊使用的武器材料"
        ],
        "ans": 0,
        "explanation": "前者著重才德，後者著重家族傳承，選人依據才是核心分別。",
        "topic": "堯舜禹"
      },
      {
        "id": "s1-hard-04",
        "q": "【情境題】一位治水者多次經過家門仍繼續工作。哪項最能說明此事獲稱許的原因？",
        "options": [
          "把公共責任置於私人生活之前",
          "藉遷居逃避原有的治理責任",
          "把治水工作交回氏族處理",
          "透過家族血緣取得共主之位"
        ],
        "ans": 0,
        "explanation": "「三過家門而不入」用以表達禹因公忘私、堅持不懈。",
        "topic": "夏朝"
      },
      {
        "id": "s1-hard-05",
        "q": "【因果題】桀、紂都承繼君位，卻施政敗壞。這最能指出世襲制哪項限制？",
        "options": [
          "血緣傳承未能確保繼任者才德",
          "君位不能在同一家族內延續",
          "每次傳位都必須重新推選共主",
          "君主施政好壞取決於都城位置"
        ],
        "ans": 0,
        "explanation": "兩個例子指向繼任者質素的問題，而不是都城或推選程序。",
        "topic": "夏商興亡"
      },
      {
        "id": "s1-hard-06",
        "q": "【證據題】若要支持「商代已有早期國家組織」，哪組資料最切合？",
        "options": [
          "設有首都、軍隊、官制及刑法",
          "以採集果實和狩獵維生",
          "以相同血緣聚居形成氏族",
          "以神話解釋文明的起源"
        ],
        "ans": 0,
        "explanation": "政治與軍事制度是國家組織的直接證據，與基本生產方式不同。",
        "topic": "商朝"
      },
      {
        "id": "s1-hard-07",
        "q": "【因果題】商紂王施酷刑、殘殺忠臣。哪條因果關係最合理？",
        "options": [
          "暴政加劇不滿，削弱統治支持",
          "酷刑增加諸侯對王室的信任",
          "殺害忠臣促使禪讓制恢復",
          "縱情享樂有助改善百姓生計"
        ],
        "ans": 0,
        "explanation": "筆記以「民心盡失」說明暴政如何削弱商朝統治。",
        "topic": "商朝"
      },
      {
        "id": "s1-hard-08",
        "q": "【政策比較】周武王封武庚，又在附近設三監，兩項措施如何互相配合？",
        "options": [
          "安撫殷民，同時防範叛亂",
          "恢復商室，同時撤走周室監察",
          "擴大殷權，同時取消地方分封",
          "遷走殷民，同時保留鎬京都城"
        ],
        "ans": 0,
        "explanation": "封武庚是安撫，設三監是監視；政策兼顧懷柔與防範。",
        "topic": "周代分封"
      },
      {
        "id": "s1-hard-09",
        "q": "【事件排序】哪個次序符合三監之亂的發展？",
        "options": [
          "成王年幼→周公攝政→三監聯合武庚作亂",
          "周公攝政→武王立國→平王東遷",
          "武庚作亂→周公攝政→成王年幼即位",
          "平王東遷→三監作亂→周武王滅商"
        ],
        "ans": 0,
        "explanation": "武王死後成王年幼，周公攝政引起三監不滿，才出現聯合叛亂。",
        "topic": "周代分封"
      },
      {
        "id": "s1-hard-10",
        "q": "【政策目的】平定三監之亂後再次分封，最能回應哪個統治需要？",
        "options": [
          "加強周室對東方地區的控制",
          "削弱周室對殷地的影響",
          "把共主之位禪讓予殷遺民",
          "使各氏族恢復獨立群居生活"
        ],
        "ans": 0,
        "explanation": "筆記把再次分封與加強東方控制聯繫起來。",
        "topic": "周代分封"
      },
      {
        "id": "s1-hard-11",
        "q": "【資料判讀】幽王寵愛褒姒、廢申后及太子；申侯聯犬戎攻周。哪項是事件間的直接關聯？",
        "options": [
          "廢后及太子引起申侯不滿，促成外部攻擊",
          "平王東遷直接引起幽王廢后",
          "三監反叛迫使申侯聯合犬戎",
          "甲骨文的出土加劇王位爭議"
        ],
        "ans": 0,
        "explanation": "題目要求依據事件找出直接因果，並非僅記人名。",
        "topic": "西周與東周"
      },
      {
        "id": "s1-hard-12",
        "q": "【因果題】鎬京遭攻陷而殘破，平王其後東遷。哪項解釋最切合筆記？",
        "options": [
          "原都城受破壞，促使周室另選都城",
          "為加強對殷民監視，周室始設三監",
          "為逃避商代水患，平王遷往殷",
          "為推行禪讓，平王放棄王畿土地"
        ],
        "ans": 0,
        "explanation": "筆記把鎬京殘破列為平王東遷洛邑的原因。",
        "topic": "西周與東周"
      },
      {
        "id": "s1-hard-13",
        "q": "【因果題】東遷後王畿縮小，周天子更難號令諸侯。哪項解釋最合理？",
        "options": [
          "直轄資源減少，削弱控制諸侯的實力",
          "直轄土地減少，使各諸侯必須更服從",
          "遷都本身令所有諸侯失去軍隊",
          "土地減少提高了天子在各國的威信"
        ],
        "ans": 0,
        "explanation": "王畿減少影響周室實力；筆記亦提及平王聲望下降。",
        "topic": "春秋戰國"
      },
      {
        "id": "s1-hard-14",
        "q": "【社會變化】戰亂促使學術向下移動，平民亦能提出政治思想。這反映甚麼變化？",
        "options": [
          "學問逐漸突破貴族壟斷",
          "學問開始只由王室掌握",
          "教育機會由平民轉向貴族",
          "各國以軍事統一取代思想交流"
        ],
        "ans": 0,
        "explanation": "關鍵是知識與教育的參與者擴大，而不是王室重新壟斷。",
        "topic": "春秋戰國"
      },
      {
        "id": "s1-hard-15",
        "q": "【概念應用】諸侯尊重周室，並協助抵禦外族，最符合哪項政治主張？",
        "options": [
          "尊王攘夷",
          "父傳子以建立家天下",
          "以疏導方法處理水患",
          "以血緣聚居形成氏族"
        ],
        "ans": 0,
        "explanation": "此情境分別對應「尊王」與「攘夷」兩部分。",
        "topic": "春秋戰國"
      }
    ]
  },
  "s2": {
    "easy": [
      {
        "id": "s2-easy-01",
        "q": "五代政權主要出現於哪個流域？",
        "options": [
          "黃河流域",
          "長江流域",
          "珠江流域",
          "閩江流域"
        ],
        "ans": 0,
        "explanation": "筆記以北方黃河流域的政權交替概括五代。",
        "topic": "五代十國"
      },
      {
        "id": "s2-easy-02",
        "q": "五代君主多由哪個職位出身？",
        "options": [
          "節度使",
          "參知政事",
          "三司使",
          "通判"
        ],
        "ans": 0,
        "explanation": "節度使掌握軍事力量，與五代武人專政有關。",
        "topic": "五代十國"
      },
      {
        "id": "s2-easy-03",
        "q": "誰為取得契丹支持而割讓燕雲十六州？",
        "options": [
          "石敬瑭",
          "趙匡胤",
          "趙構",
          "王安石"
        ],
        "ans": 0,
        "explanation": "筆記把割讓燕雲十六州一事歸於石敬瑭。",
        "topic": "五代十國"
      },
      {
        "id": "s2-easy-04",
        "q": "北宋開國君主趙匡胤的廟號是甚麼？",
        "options": [
          "宋太祖",
          "宋神宗",
          "宋徽宗",
          "宋高宗"
        ],
        "ans": 0,
        "explanation": "趙匡胤建立北宋，稱宋太祖。",
        "topic": "宋初集權"
      },
      {
        "id": "s2-easy-05",
        "q": "趙匡胤在哪場兵變中「黃袍加身」？",
        "options": [
          "陳橋兵變",
          "靖康之難",
          "三監之亂",
          "安史之亂"
        ],
        "ans": 0,
        "explanation": "陳橋兵變與宋太祖建立宋朝直接相關。",
        "topic": "宋初集權"
      },
      {
        "id": "s2-easy-06",
        "q": "北宋的首都是哪裏？",
        "options": [
          "汴京",
          "臨安",
          "洛邑",
          "鎬京"
        ],
        "ans": 0,
        "explanation": "北宋定都汴京；南宋定都臨安。",
        "topic": "宋初集權"
      },
      {
        "id": "s2-easy-07",
        "q": "宋太祖以哪件事收回禁軍將領兵權？",
        "options": [
          "杯酒釋兵權",
          "陳橋兵變",
          "靖康之難",
          "澶淵之盟"
        ],
        "ans": 0,
        "explanation": "杯酒釋兵權是宋太祖削弱禁軍將領權力的措施。",
        "topic": "宋初集權"
      },
      {
        "id": "s2-easy-08",
        "q": "宋代的中央精銳軍隊稱為甚麼？",
        "options": [
          "禁軍",
          "廂軍",
          "保甲",
          "地方鄉兵"
        ],
        "ans": 0,
        "explanation": "筆記區分中央禁軍與地方廂軍。",
        "topic": "宋初集權"
      },
      {
        "id": "s2-easy-09",
        "q": "宋神宗任用誰推行熙寧變法？",
        "options": [
          "王安石",
          "司馬光",
          "秦檜",
          "韓世忠"
        ],
        "ans": 0,
        "explanation": "熙寧變法由宋神宗支持、王安石推行。",
        "topic": "北宋變法"
      },
      {
        "id": "s2-easy-10",
        "q": "熙寧變法的主要目標是甚麼？",
        "options": [
          "富國強兵",
          "重文輕武",
          "尊王攘夷",
          "強榦弱枝"
        ],
        "ans": 0,
        "explanation": "富國與強兵是變法目標；其餘是其他背景下的主張。",
        "topic": "北宋變法"
      },
      {
        "id": "s2-easy-11",
        "q": "靖康之難中，哪個政權擄走徽、欽二帝？",
        "options": [
          "金",
          "遼",
          "蒙古",
          "西夏"
        ],
        "ans": 0,
        "explanation": "金滅遼後進攻北宋，擄走徽、欽二帝。",
        "topic": "北宋覆亡"
      },
      {
        "id": "s2-easy-12",
        "q": "誰建立南宋？",
        "options": [
          "宋高宗趙構",
          "宋太祖趙匡胤",
          "宋神宗",
          "宋欽宗"
        ],
        "ans": 0,
        "explanation": "趙構南渡建立南宋，是為宋高宗。",
        "topic": "南宋興亡"
      },
      {
        "id": "s2-easy-13",
        "q": "南宋以哪裏為首都？",
        "options": [
          "臨安",
          "汴京",
          "洛邑",
          "長安"
        ],
        "ans": 0,
        "explanation": "臨安即今杭州，與北宋汴京須加區分。",
        "topic": "南宋興亡"
      },
      {
        "id": "s2-easy-14",
        "q": "北宋中期出現、中國最早的紙幣稱為甚麼？",
        "options": [
          "交子",
          "銅錢",
          "歲幣",
          "絲帛"
        ],
        "ans": 0,
        "explanation": "交子是紙幣；銅錢是鑄幣，歲幣是支付給其他政權的財物。",
        "topic": "宋代經濟"
      },
      {
        "id": "s2-easy-15",
        "q": "宋代哪個機構專責管理海外貿易？",
        "options": [
          "市舶司",
          "三司",
          "樞密院",
          "地方通判"
        ],
        "ans": 0,
        "explanation": "市舶司負責海外貿易；三司、樞密及通判各有不同職能。",
        "topic": "海外貿易"
      }
    ],
    "medium": [
      {
        "id": "s2-medium-01",
        "q": "五代與十國的政權形勢有何主要分別？",
        "options": [
          "五代在北方交替，十國多在南方並存",
          "五代在北方並存，十國在南方交替",
          "五代在南方交替，十國多在北方並存",
          "五代在南方並存，十國在北方交替"
        ],
        "ans": 0,
        "explanation": "題目比較的是地理分布與政權交替、並存的方式。",
        "topic": "五代十國"
      },
      {
        "id": "s2-medium-02",
        "q": "依據筆記，南方十國為何相對較穩定？",
        "options": [
          "君主較重視民生，政局經濟較穩定",
          "君主均出自禁軍，經常發動兵變",
          "君主以輕視文治維持社會秩序",
          "君主透過割讓燕雲十六州安定南方"
        ],
        "ans": 0,
        "explanation": "筆記把南方較重民生，與北方武人專政的局面作對照。",
        "topic": "五代十國"
      },
      {
        "id": "s2-medium-03",
        "q": "失去燕雲十六州對北宋國防有甚麼影響？",
        "options": [
          "失去長城屏障，北方較難防守",
          "失去長江天險，南方較難防守",
          "失去汴京都城，北宋立即南渡",
          "失去市舶收入，海上貿易停頓"
        ],
        "ans": 0,
        "explanation": "長城屏障及北方平原地勢，是筆記解釋國防危機的重點。",
        "topic": "五代十國"
      },
      {
        "id": "s2-medium-04",
        "q": "宋代以文人掌軍，主要反映甚麼政策？",
        "options": [
          "重文輕武",
          "尊王攘夷",
          "聯金滅遼",
          "聯蒙滅金"
        ],
        "ans": 0,
        "explanation": "以文人掌軍與重視文治、抑制武人勢力的方向一致。",
        "topic": "宋初集權"
      },
      {
        "id": "s2-medium-05",
        "q": "宋代以哪個官職作副相，分割相權？",
        "options": [
          "參知政事",
          "樞密使",
          "三司使",
          "通判"
        ],
        "ans": 0,
        "explanation": "參知政事為副相；樞密使分軍權，三司使分財權。",
        "topic": "宋初集權"
      },
      {
        "id": "s2-medium-06",
        "q": "宋代以哪個官職分割宰相的軍權？",
        "options": [
          "樞密使",
          "參知政事",
          "三司使",
          "通判"
        ],
        "ans": 0,
        "explanation": "樞密使負責軍事，與財政及地方監察不同。",
        "topic": "宋初集權"
      },
      {
        "id": "s2-medium-07",
        "q": "宋代以哪個官職分割宰相的財權？",
        "options": [
          "三司使",
          "樞密使",
          "參知政事",
          "知州"
        ],
        "ans": 0,
        "explanation": "三司使分割財權；樞密使分割軍權。",
        "topic": "宋初集權"
      },
      {
        "id": "s2-medium-08",
        "q": "宋代在地方設通判，主要目的為何？",
        "options": [
          "監察知州，牽制地方權力",
          "代替樞密使統領中央軍事",
          "代替三司使管理中央財政",
          "代替市舶司管理海外貿易"
        ],
        "ans": 0,
        "explanation": "通判的監察對象是地方知州。",
        "topic": "宋初集權"
      },
      {
        "id": "s2-medium-09",
        "q": "宋代更戍法的主要安排是甚麼？",
        "options": [
          "軍隊三年一調防",
          "固定將領長期訓練同一部隊",
          "地方軍隊改由市舶司統領",
          "軍隊按土地面積分配稅額"
        ],
        "ans": 0,
        "explanation": "更戍法導致兵將分離；置將法則著重固定兵將。",
        "topic": "宋初集權"
      },
      {
        "id": "s2-medium-10",
        "q": "宋與遼訂立哪個和約，以歲幣換取和平？",
        "options": [
          "澶淵之盟",
          "陳橋兵變",
          "熙寧變法",
          "靖康之難"
        ],
        "ans": 0,
        "explanation": "澶淵之盟是和約；其餘是兵變、變法及亡國事件。",
        "topic": "宋初集權"
      },
      {
        "id": "s2-medium-11",
        "q": "王安石哪項新法固定兵將，以提高默契？",
        "options": [
          "置將法",
          "更戍法",
          "保甲法",
          "方田均稅法"
        ],
        "ans": 0,
        "explanation": "置將法針對更戍法兵將分離的問題。",
        "topic": "北宋變法"
      },
      {
        "id": "s2-medium-12",
        "q": "哪位官員反對王安石新法？",
        "options": [
          "司馬光",
          "宋神宗",
          "宋太祖",
          "韓世忠"
        ],
        "ans": 0,
        "explanation": "筆記把司馬光列為反對新法的保守派官員。",
        "topic": "北宋變法"
      },
      {
        "id": "s2-medium-13",
        "q": "南宋主要依靠哪條河流天險抵抗金人？",
        "options": [
          "長江",
          "黃河",
          "汴河",
          "珠江"
        ],
        "ans": 0,
        "explanation": "長江有助限制金人騎兵優勢，支持南宋偏安。",
        "topic": "南宋興亡"
      },
      {
        "id": "s2-medium-14",
        "q": "哪項人物與戰役的配對正確？",
        "options": [
          "韓世忠—黃天蕩之役",
          "岳飛—陳橋兵變",
          "趙匡胤—靖康之難",
          "王安石—黃天蕩之役"
        ],
        "ans": 0,
        "explanation": "韓世忠在黃天蕩之役圍困金兵。",
        "topic": "南宋興亡"
      },
      {
        "id": "s2-medium-15",
        "q": "哪項最能作為北宋汴京城市生活的圖像材料？",
        "options": [
          "《清明上河圖》",
          "記錄商代活動的甲骨文",
          "宋與遼訂立的和約",
          "宋代地方官員職權表"
        ],
        "ans": 0,
        "explanation": "《清明上河圖》反映汴京繁華，與制度表或和約的用途不同。",
        "topic": "宋代經濟"
      }
    ],
    "hard": [
      {
        "id": "s2-hard-01",
        "q": "【政策比較】中央集中精銳，地方留下較弱軍隊，最能體現哪個政策方向？",
        "options": [
          "強榦弱枝，以中央制約地方",
          "以地方兵力制約中央禁軍",
          "透過市舶收入取代中央軍隊",
          "以固定兵將取代地方監察"
        ],
        "ans": 0,
        "explanation": "「榦」指中央，「枝」指地方；資源配置使中央相對強大。",
        "topic": "宋初集權"
      },
      {
        "id": "s2-hard-02",
        "q": "【兩面影響】「強榦弱枝」一方面防地方挑戰中央，另一方面造成甚麼問題？",
        "options": [
          "地方兵力偏弱，難抗外族進攻",
          "地方兵力偏強，中央無兵可用",
          "地方知州改為統領樞密院",
          "地方不再需要承擔任何防務"
        ],
        "ans": 0,
        "explanation": "政策同時有鞏固中央與削弱地方防衛的效果。",
        "topic": "宋初集權"
      },
      {
        "id": "s2-hard-03",
        "q": "【因果題】更戍法削弱戰鬥力的推理，哪項最完整？",
        "options": [
          "調防使兵將分離→默契不足→戰鬥力下降",
          "固定兵將→訓練熟悉→戰鬥力下降",
          "增加稅收→地方裁兵→兵將更熟悉",
          "以文人掌軍→廢除禁軍→各地失守"
        ],
        "ans": 0,
        "explanation": "關鍵連繫是兵將分離及缺乏默契，不是泛指文人掌軍。",
        "topic": "宋初集權"
      },
      {
        "id": "s2-hard-04",
        "q": "【兩面影響】多設官職互相監察，除分散權力外，可能產生甚麼弊端？",
        "options": [
          "政出多門，降低行政效能",
          "相權集中，官員失去監察作用",
          "地方知州得以不受中央牽制",
          "軍隊固定配屬將領而免調防"
        ],
        "ans": 0,
        "explanation": "互相牽制有集權作用，但職權重疊亦可降低效率。",
        "topic": "宋初集權"
      },
      {
        "id": "s2-hard-05",
        "q": "【綜合題】冗員、冗兵及歲幣三者的共同影響是甚麼？",
        "options": [
          "擴大支出，加重財政負擔",
          "減少支出，消除財政積貧",
          "提高默契，增強軍隊戰鬥力",
          "擴大王畿，增加直轄土地"
        ],
        "ans": 0,
        "explanation": "三者都可增加開支，連結宋代冗費及積貧問題。",
        "topic": "宋初集權"
      },
      {
        "id": "s2-hard-06",
        "q": "【政策配對】哪組新法及作用配對最恰當？",
        "options": [
          "農田水利法、方田均稅法—促進農業及稅收",
          "置將法、更戍法—共同推行兵將分離",
          "保甲法、方田均稅法—共同分割宰相軍權",
          "農田水利法、置將法—共同監察地方知州"
        ],
        "ans": 0,
        "explanation": "前兩項屬農業及財政措施；置將法則針對軍事問題。",
        "topic": "北宋變法"
      },
      {
        "id": "s2-hard-07",
        "q": "【因果題】新法損害部分官員及豪強利益，如何影響變法？",
        "options": [
          "引起反對及新舊黨爭，妨礙推行",
          "使原有利益群體一致支持新法",
          "使外族立即停止進攻宋境",
          "令更戍法與置將法成為同一制度"
        ],
        "ans": 0,
        "explanation": "利益受損引發政治反對，是筆記解釋變法失敗的因素之一。",
        "topic": "北宋變法"
      },
      {
        "id": "s2-hard-08",
        "q": "【事件分析】「聯金滅遼」後金再攻宋，最能反映哪個外交風險？",
        "options": [
          "盟友擊敗共同敵人後，可能轉而威脅自己",
          "共同敵人滅亡後，盟友必會歸還全部領土",
          "簽訂軍事聯盟能確保長期和平",
          "地理屏障可完全取代軍事部署"
        ],
        "ans": 0,
        "explanation": "此題由事件推論聯盟風險，並不把所有聯盟視為必然失敗。",
        "topic": "北宋覆亡"
      },
      {
        "id": "s2-hard-09",
        "q": "【經濟因果】北方人口南遷如何幫助南宋維持政權？",
        "options": [
          "人力技術資源增加→經濟發展→稅收支持政權",
          "人口南遷→北方農業增產→南宋軍費下降",
          "人口南遷→地方停止生產→海外需求增加",
          "技術南傳→中央撤銷稅收→軍費完全免除"
        ],
        "ans": 0,
        "explanation": "筆記把人口資源南移、南方經濟及支持南宋的稅收連結起來。",
        "topic": "南宋興亡"
      },
      {
        "id": "s2-hard-10",
        "q": "【地理應用】金軍擅長騎兵，南宋利用長江防守。長江的主要作用是甚麼？",
        "options": [
          "形成渡江障礙，限制騎兵優勢",
          "提供北方平原，便利騎兵展開",
          "使金軍更容易沿陸路直抵臨安",
          "以商業收入直接取代前線軍隊"
        ],
        "ans": 0,
        "explanation": "河流天險能改變作戰條件，但不代表南宋不用軍隊。",
        "topic": "南宋興亡"
      },
      {
        "id": "s2-hard-11",
        "q": "【綜合題】哪組因素最能解釋南宋得以偏安？",
        "options": [
          "長江天險、南方經濟及良將抵抗",
          "燕雲屏障、取消稅收及地方無兵",
          "北方統一、廢除海貿及文人退政",
          "甲骨文字、禪讓制度及三監監察"
        ],
        "ans": 0,
        "explanation": "筆記從地理、經濟及軍事三方面解釋南宋偏安。",
        "topic": "南宋興亡"
      },
      {
        "id": "s2-hard-12",
        "q": "【比較題】北宋聯金滅遼與南宋聯蒙滅金，有何共同後續發展？",
        "options": [
          "盟友消滅共同敵人後，轉而攻宋",
          "宋室兩次均收復全部北方領土",
          "共同敵人兩次均與宋聯手滅盟友",
          "兩次聯盟均直接促成宋代紙幣出現"
        ],
        "ans": 0,
        "explanation": "兩段事件都呈現借助新興勢力後再受其威脅的情況。",
        "topic": "南宋興亡"
      },
      {
        "id": "s2-hard-13",
        "q": "【制度影響】宋代打破唐代坊市制規限，為何有助商業？",
        "options": [
          "減少原有交易規限，增加商業活動空間",
          "把商業集中於更少的固定地點",
          "使市舶司代替全部地方市場",
          "把所有城市改為中央軍事營地"
        ],
        "ans": 0,
        "explanation": "本題考規限與商業發展的關係，不要求筆記以外的細節。",
        "topic": "宋代經濟"
      },
      {
        "id": "s2-hard-14",
        "q": "【因果題】陸上貿易受阻，但海外商品需求仍大，南宋較可能採取甚麼方向？",
        "options": [
          "發展海上貿易，開拓交易及稅收來源",
          "停止對外交易，把市舶司改為軍隊",
          "只增加北方陸路運輸，撤去海港管理",
          "以歲幣代替所有商人對外交易"
        ],
        "ans": 0,
        "explanation": "筆記把陸路受阻、海外需求及海貿稅收相連結。",
        "topic": "海外貿易"
      },
      {
        "id": "s2-hard-15",
        "q": "【資料判讀】某港大量出口茶葉、絲綢、瓷器，官府設機構管理。最合理的配對是甚麼？",
        "options": [
          "海外貿易—市舶司",
          "地方監察—通判",
          "中央軍事—樞密使",
          "宰相副職—參知政事"
        ],
        "ans": 0,
        "explanation": "出口商品與港口管理指向海外貿易，而非一般財政或軍事職能。",
        "topic": "海外貿易"
      }
    ]
  }
}; }
