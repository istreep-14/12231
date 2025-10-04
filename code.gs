// ============================================
// CONFIGURATION
// ============================================
const CONFIG = {
  USERNAME: 'frankscobey',
  MAX_GAMES_PER_BATCH: 3,
  AUTO_FETCH_CALLBACK_DATA: true // Automatically fetch callback data for new games
};

const SHEETS = {
  GAMES: 'Games',
  ANALYSIS: 'Analysis',
  CALLBACK: 'Callback',
  DERIVED: 'Derived Data',
  OPENINGS_DB: 'Openings DB' // TSV-backed openings database (Name, Trim Slug, Family, Name, Variation1..6)
};

// Result to outcome mapping
const RESULT_MAP = {
  'win': 'Win',
  'checkmated': 'Loss',
  'agreed': 'Draw',
  'repetition': 'Draw',
  'timeout': 'Loss',
  'resigned': 'Loss',
  'stalemate': 'Draw',
  'lose': 'Loss',
  'insufficient': 'Draw',
  '50move': 'Draw',
  'abandoned': 'Loss',
  'kingofthehill': 'Loss',
  'threecheck': 'Loss',
  'timevsinsufficient': 'Draw',
  'bughousepartnerlose': 'Loss'
};

// Termination descriptions
const TERMINATION_MAP = {
  'win': 'Win',
  'checkmated': 'Checkmate',
  'agreed': 'Agreement',
  'repetition': 'Repetition',
  'timeout': 'Timeout',
  'resigned': 'Resignation',
  'stalemate': 'Stalemate',
  'lose': 'Loss',
  'insufficient': 'Insufficient material',
  '50move': '50-move rule',
  'abandoned': 'Abandoned',
  'kingofthehill': 'King of the Hill',
  'threecheck': 'Three-check',
  'timevsinsufficient': 'Timeout vs insufficient',
  'bughousepartnerlose': 'Bughouse partner lost'
};

// Cache for openings database (sheet-backed)
let OPENINGS_DB_CACHE = null;
const OPENINGS_DB_HEADERS = [
  'Name', 'Trim Slug', 'Family', 'Name',
  'Variation 1', 'Variation 2', 'Variation 3', 'Variation 4', 'Variation 5', 'Variation 6'
];

// Minimal opening outputs to store in-sheet
const DERIVED_OPENING_HEADERS = [
  'Opening Name', 'Opening Family'
];

// ============================================
// MAIN MENU
// ============================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('♟️ Chess Analyzer')
    .addItem('1️⃣ Setup Sheets', 'setupSheets')
    .addItem('2️⃣ Initial Fetch (All Games)', 'fetchAllGamesInitial')
    .addItem('3️⃣ Update Recent Games', 'fetchChesscomGames')
    .addSeparator()
    .addItem('📋 Fetch Callback Last 10', 'fetchCallbackLast10')
    .addItem('🔗 Refresh Opening Mappings', 'refreshDerivedDbMappings')
    .addItem('🧪 Enrich Move Times (Selection)', 'enrichMoveTimesForSelection')
    .addToUi();
}


// ============================================
// CALLBACK DATA FETCHING
// ============================================


function fetchCallbackData(game) {
  // Validate game object has required fields
  if (!game || !game.gameId || !game.timeClass || !game.white || !game.black) {
    Logger.log(`Skipping callback fetch - incomplete game data: ${JSON.stringify(game)}`);
    return null;
  }
  
  const gameId = game.gameId;
  const timeClass = game.timeClass.toLowerCase();
  const gameType = timeClass === 'daily' ? 'daily' : 'live';
  const callbackUrl = `https://www.chess.com/callback/${gameType}/game/${gameId}`;
  
  try {
    const response = UrlFetchApp.fetch(callbackUrl, {muteHttpExceptions: true});
    
    if (response.getResponseCode() !== 200) {
      Logger.log(`Callback API error for game ${gameId}: ${response.getResponseCode()}`);
      return null;
    }
    
    const data = JSON.parse(response.getContentText());
    
    if (!data || !data.game) {
      Logger.log(`Invalid callback data for game ${gameId}`);
      return null;
    }
    
    const gameData = data.game;
    const players = data.players || {};
    const topPlayer = players.top || {};
    const bottomPlayer = players.bottom || {};
    
    // Determine my color and player data
    const isWhite = game.white.toLowerCase() === CONFIG.USERNAME.toLowerCase();
    const myColor = isWhite ? 'white' : 'black';
    
    let myRatingChange = isWhite ? gameData.ratingChangeWhite : gameData.ratingChangeBlack;
    let oppRatingChange = isWhite ? gameData.ratingChangeBlack : gameData.ratingChangeWhite;
    
    // If rating change is 0, it's likely an error (unless edge case draw)
    // Set to null to indicate unreliable data
    if (myRatingChange === 0) myRatingChange = null;
    if (oppRatingChange === 0) oppRatingChange = null;
    
    // Get player data (top/bottom can be either color)
    let whitePlayer, blackPlayer;
    if (topPlayer.color === 'white') {
      whitePlayer = topPlayer;
      blackPlayer = bottomPlayer;
    } else {
      whitePlayer = bottomPlayer;
      blackPlayer = topPlayer;
    }
    
    // Determine my player and opponent player
    const myPlayer = isWhite ? whitePlayer : blackPlayer;
    const oppPlayer = isWhite ? blackPlayer : whitePlayer;
    
    // Get ratings from callback
    const myRating = myPlayer.rating || null;
    const oppRating = oppPlayer.rating || null;
    
    // Calculate "before" ratings by subtracting rating change
    let myRatingBefore = null;
    let oppRatingBefore = null;
    
    if (myRating !== null && myRatingChange !== null) {
      myRatingBefore = myRating - myRatingChange;
    }
    if (oppRating !== null && oppRatingChange !== null) {
      oppRatingBefore = oppRating - oppRatingChange;
    }
    
    return {
      gameId: gameId,
      gameUrl: game.gameUrl,
      callbackUrl: callbackUrl,
      endTime: gameData.endTime,
      myColor: myColor,
      timeClass: game.timeClass,
      myRating: myRating,
      oppRating: oppRating,
      myRatingChange: myRatingChange,
      oppRatingChange: oppRatingChange,
      myRatingBefore: myRatingBefore,
      oppRatingBefore: oppRatingBefore,
      baseTime: gameData.baseTime1 || 0,
      timeIncrement: gameData.timeIncrement1 || 0,
      moveTimestamps: gameData.moveTimestamps ? String(gameData.moveTimestamps) : '',
      myUsername: myPlayer.username || '',
      myCountry: myPlayer.countryName || '',
      myMembership: myPlayer.membershipCode || '',
      myMemberSince: myPlayer.memberSince || 0,
      myDefaultTab: myPlayer.defaultTab || null,
      myPostMoveAction: myPlayer.postMoveAction || '',
      myLocation: myPlayer.location || '',
      oppUsername: oppPlayer.username || '',
      oppCountry: oppPlayer.countryName || '',
      oppMembership: oppPlayer.membershipCode || '',
      oppMemberSince: oppPlayer.memberSince || 0,
      oppDefaultTab: oppPlayer.defaultTab || null,
      oppPostMoveAction: oppPlayer.postMoveAction || '',
      oppLocation: oppPlayer.location || ''
    };
    
  } catch (error) {
    Logger.log(`Error fetching callback data for game ${gameId}: ${error.message}`);
    return null;
  }
}

function saveCallbackData(callbackData) {
  if (!callbackData) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const callbackSheet = ss.getSheetByName(SHEETS.CALLBACK);
  
  if (!callbackSheet) return;
  
  const row = [
    callbackData.gameId,
    callbackData.gameUrl,
    callbackData.callbackUrl,
    callbackData.endTime,
    callbackData.myColor,
    callbackData.timeClass,
    callbackData.myRating,
    callbackData.oppRating,
    callbackData.myRatingChange,
    callbackData.oppRatingChange,
    callbackData.myRatingBefore,
    callbackData.oppRatingBefore,
    callbackData.baseTime,
    callbackData.timeIncrement,
    callbackData.moveTimestamps,
    callbackData.myUsername,
    callbackData.myCountry,
    callbackData.myMembership,
    callbackData.myMemberSince,
    callbackData.myDefaultTab,
    callbackData.myPostMoveAction,
    callbackData.myLocation,
    callbackData.oppUsername,
    callbackData.oppCountry,
    callbackData.oppMembership,
    callbackData.oppMemberSince,
    callbackData.oppDefaultTab,
    callbackData.oppPostMoveAction,
    callbackData.oppLocation,
    new Date()
  ];
  
  const lastRow = callbackSheet.getLastRow();
  callbackSheet.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);
}

function processNewGamesAutoFeatures(newGames) {
  if (!newGames || newGames.length === 0) return;
  
  // Auto-fetch callback data
  if (CONFIG.AUTO_FETCH_CALLBACK_DATA && newGames.length <= CONFIG.MAX_GAMES_PER_BATCH) {
    fetchCallbackForGames(newGames);
  }
  
  // Auto-analyze new games
  if (CONFIG.AUTO_ANALYZE_NEW_GAMES && newGames.length <= CONFIG.MAX_GAMES_PER_BATCH) {
    analyzeGames(newGames);
  }
}

// ============================================
// FETCH CALLBACK DATA FOR GAMES
// ============================================
function fetchCallbackLast10() { fetchCallbackLastN(10); }
function fetchCallbackLast25() { fetchCallbackLastN(25); }
function fetchCallbackLast50() { fetchCallbackLastN(50); }

function fetchCallbackLastN(count) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const callbackSheet = ss.getSheetByName(SHEETS.CALLBACK);
  
  if (!gamesSheet || !callbackSheet) {
    SpreadsheetApp.getUi().alert('❌ Please run "Setup Sheets" first!');
    return;
  }
  
  const gamesWithoutCallback = getGamesWithoutCallback(count);
  
  if (gamesWithoutCallback.length === 0) {
    SpreadsheetApp.getUi().alert('✅ No games need callback data!');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    `Fetch callback data for ${gamesWithoutCallback.length} game(s)?`,
    `This will fetch detailed game data from Chess.com.\n\nContinue?`,
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  fetchCallbackForGames(gamesWithoutCallback);
}

function getGamesWithoutCallback(maxCount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const data = gamesSheet.getDataRange().getValues();
  const games = [];
  
  // Iterate from newest to oldest (reverse order)
  for (let i = data.length - 1; i >= 1 && games.length < maxCount; i--) {
    if (data[i][13] === false) { // Callback Fetched column (index 13)
      const myColor = data[i][3];
      const opponent = data[i][4];
      const gameId = data[i][11];
      
      // Get time class from derived data
      const derivedSheet = ss.getSheetByName(SHEETS.DERIVED);
      let timeClass = '';
      
      if (derivedSheet) {
        const derivedData = derivedSheet.getDataRange().getValues();
        for (let j = 1; j < derivedData.length; j++) {
          if (derivedData[j][0] === gameId) {
            timeClass = derivedData[j][5];
            break;
          }
        }
      }
      
      games.push({
        row: i + 1,
        gameId: gameId,
        gameUrl: data[i][0],
        white: myColor === 'white' ? CONFIG.USERNAME : opponent,
        black: myColor === 'black' ? CONFIG.USERNAME : opponent,
        timeClass: timeClass
      });
    }
  }
  
  return games.reverse(); // Return in chronological order (oldest first)
}

function fetchCallbackForGames(gamesToFetch) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  
  let successCount = 0;
  let errorCount = 0;
  
  ss.toast('Fetching callback data...', '📋', -1);
  
  for (let i = 0; i < gamesToFetch.length; i++) {
    const game = gamesToFetch[i];
    
    try {
      ss.toast(`Fetching callback ${i + 1} of ${gamesToFetch.length}...`, '📋', -1);
      
      const callbackData = fetchCallbackData(game);
      if (callbackData) {
        saveCallbackData(callbackData);
        
        // Mark callback as fetched
        if (game.row) {
          gamesSheet.getRange(game.row, 14).setValue(true); // Callback Fetched column (index 14)
        }
        
        successCount++;
      } else {
        errorCount++;
      }
      
      Utilities.sleep(300); // Rate limiting
      
    } catch (error) {
      Logger.log(`Error fetching callback for game ${game.gameId}: ${error}`);
      errorCount++;
    }
  }
  
  ss.toast(`✅ Callback fetched: ${successCount}, Errors: ${errorCount}`, '📋', 5);
  
}

function findGameRow(gameId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const data = gamesSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][11] === gameId) { // Game ID column (index 11)
      return i + 1;
    }
  }
  return -1;
}
// ============================================
// FETCH CHESS.COM GAMES
// ============================================

// INITIAL FETCH: Get all games from all archives
function fetchAllGamesInitial() {
  const username = CONFIG.USERNAME;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  
  if (!gamesSheet) {
    SpreadsheetApp.getUi().alert('❌ Please run "Setup Sheets" first!');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Initial Full Fetch',
    'This will fetch ALL games from your Chess.com history.\n' +
    'This may take several minutes depending on how many games you have.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    ss.toast('Fetching all game archives...', '⏳', -1);
    
    const archivesUrl = `https://api.chess.com/pub/player/${username}/games/archives`;
    const archivesResponse = UrlFetchApp.fetch(archivesUrl);
    const archives = JSON.parse(archivesResponse.getContentText()).archives;
    
    ss.toast(`Found ${archives.length} archives. Fetching games...`, '⏳', -1);
    
    const allGames = [];
    const props = PropertiesService.getScriptProperties();
    
    const now = new Date();
    const currentYearMonth = `${now.getFullYear()}_${String(now.getMonth() + 1).padStart(2, '0')}`;
    
    for (let i = 0; i < archives.length; i++) {
      ss.toast(`Fetching archive ${i + 1} of ${archives.length}...`, '⏳', -1);
      Utilities.sleep(500);
      
      const response = fetchWithETag(archives[i], null);
      if (response.data) {
        allGames.push(...response.data.games);
        
        // Only store ETag if this is the current month (mutable archive)
        const archiveKey = archiveUrlToKey(archives[i]);
        if (archiveKey === currentYearMonth) {
          props.setProperty('etag_current', response.etag);
        }
        // Don't store ETags for completed months - they're immutable
      }
      
      Logger.log(`Archive ${i + 1}/${archives.length}: ${response.data?.games?.length || 0} games`);
    }
    
    // Filter out duplicates before processing
    const existingGameIds = new Set();
    if (gamesSheet.getLastRow() > 1) {
      const existingData = gamesSheet.getDataRange().getValues();
      for (let i = 1; i < existingData.length; i++) {
        existingGameIds.add(existingData[i][11]); // Game ID column (index 11)
      }
    }
    
    const newGames = allGames.filter(game => {
      const gameId = game.url.split('/').pop();
      return !existingGameIds.has(gameId);
    });
    
    ss.toast(`Processing ${newGames.length} games...`, '⏳', -1);
    const rows = processGamesData(newGames, username);
    
    if (rows.length > 0) {
      const lastRow = gamesSheet.getLastRow();
      gamesSheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
      
      // Find and store the most recent game URL (latest end_time)
      let mostRecentGame = newGames[0];
      for (const game of newGames) {
        if (game.end_time > mostRecentGame.end_time) {
          mostRecentGame = game;
        }
      }
      props.setProperty('LAST_GAME_URL', mostRecentGame.url);
      props.setProperty('INITIAL_FETCH_COMPLETE', 'true');
      // Mark previous month as finalized, since initial fetch pulled all archives
      const prevMonthDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
      const prevYear = prevMonthDate.getFullYear();
      const prevMonth = String(prevMonthDate.getMonth() + 1).padStart(2, '0');
      const prevYearMonth = `${prevYear}_${prevMonth}`;
      if (prevYearMonth !== currentYearMonth) {
        props.setProperty('LAST_FINALIZED_MONTH', prevYearMonth);
      }
      
      ss.toast(`✅ Fetched ${newGames.length} games!`, '✅', 5);
      
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Error: ${error.message}`);
    Logger.log(error);
  }
}

// UPDATE FETCH: Check most recent archive(s) for new games
function fetchChesscomGames() {
  const username = CONFIG.USERNAME;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  
  if (!gamesSheet) {
    SpreadsheetApp.getUi().alert('❌ Please run "Setup Sheets" first!');
    return;
  }
  
  const props = PropertiesService.getScriptProperties();
  const initialFetchComplete = props.getProperty('INITIAL_FETCH_COMPLETE');
  
  if (!initialFetchComplete) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'No Initial Fetch Detected',
      'It looks like you haven\'t done an initial full fetch yet.\n\n' +
      'Would you like to do a quick recent fetch (current month)?\n',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) return;
  }
  
  try {
    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = String(now.getMonth() + 1).padStart(2, '0');
    const currentYearMonth = `${currentYear}_${currentMonth}`;
    
    // Calculate current month archive URL directly (no API call needed)
    const currentArchiveUrl = `https://api.chess.com/pub/player/${username}/games/${currentYear}/${currentMonth}`;
    
    const archivesToCheck = [];
    const lastKnownGameUrl = props.getProperty('LAST_GAME_URL');
    const allGames = [];
    let foundLastKnownGame = false;
    
    // Check if we need to finalize previous month
    // If last_finalized_month is not the previous month, we should fetch it once
    const prevMonthDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const prevYear = prevMonthDate.getFullYear();
    const prevMonth = String(prevMonthDate.getMonth() + 1).padStart(2, '0');
    const prevYearMonth = `${prevYear}_${prevMonth}`;
    const lastFinalizedMonth = props.getProperty('LAST_FINALIZED_MONTH') || '';

    // If previous month hasn't been finalized and it's not current month, fetch it
    if (lastFinalizedMonth !== prevYearMonth && prevYearMonth !== currentYearMonth) {
      // If Games sheet already contains any row from the previous month, consider it finalized
      if (isMonthPresentInGamesSheet(prevYear, parseInt(prevMonth, 10))) {
        props.setProperty('LAST_FINALIZED_MONTH', prevYearMonth);
      } else {
        const prevArchiveUrl = `https://api.chess.com/pub/player/${username}/games/${prevYear}/${prevMonth}`;
        archivesToCheck.push({url: prevArchiveUrl, key: prevYearMonth, isCurrent: false});
        ss.toast('Finalizing previous month...', '⏳', -1);
      }
    }
    
    // Always check current month
    archivesToCheck.push({url: currentArchiveUrl, key: currentYearMonth, isCurrent: true});
    
    for (const archive of archivesToCheck) {
      Utilities.sleep(500);
      
      // Only use ETag for current month
      const storedETag = archive.isCurrent ? props.getProperty('etag_current') : null;
      
      const response = fetchWithETag(archive.url, storedETag);
      
      if (!response.data) {
        Logger.log(`Archive ${archive.url} not modified (ETag match)`);
        continue;
      }
      
      // Only store ETag if this is current month
      if (archive.isCurrent) {
        props.setProperty('etag_current', response.etag);
      } else {
        // Mark previous month as finalized
        props.setProperty('LAST_FINALIZED_MONTH', archive.key);
      }
      
      const gamesData = (response.data && response.data.games) ? response.data.games : [];
      
      if (lastKnownGameUrl) {
        // Add only games strictly newer than lastKnownGameUrl
        let encounteredLast = false;
        for (let i = 0; i < gamesData.length; i++) {
          const game = gamesData[i];
          if (game.url === lastKnownGameUrl) {
            encounteredLast = true;
            foundLastKnownGame = true;
            break;
          }
        }
        if (!encounteredLast) {
          // Whole archive is newer; append all
          allGames.push(...gamesData);
        } else {
          // Collect only newer ones (after lastKnown)
          for (let i = gamesData.length - 1; i >= 0; i--) {
            const game = gamesData[i];
            if (game.url === lastKnownGameUrl) break;
            allGames.unshift(game);
          }
        }
      } else {
        // No last known; but to avoid refetching entire previous month, rely on ETag check + duplicate filter.
        allGames.push(...gamesData);
      }
    }
    
    if (allGames.length === 0) {
      ss.toast('No new games found!', 'ℹ️', 3);
      return;
    }
    
    // Filter out duplicates before processing
    // Also skip any game older than or equal to last known by end_time if URL unknown
    const existingGameIds = new Set();
    let lastKnownEndTime = 0;
    if (gamesSheet.getLastRow() > 1) {
      const existingData = gamesSheet.getDataRange().getValues();
      for (let i = 1; i < existingData.length; i++) {
        existingGameIds.add(existingData[i][11]); // Game ID column (index 11)
        const endDate = existingData[i][1];
        const endTime = existingData[i][2];
        if (endDate instanceof Date && endTime instanceof Date) {
          const ts = new Date(
            endDate.getFullYear(), endDate.getMonth(), endDate.getDate(),
            endTime.getHours(), endTime.getMinutes(), endTime.getSeconds()
          ).getTime() / 1000;
          if (ts > lastKnownEndTime) lastKnownEndTime = ts;
        }
      }
    }

    const newGames = allGames.filter(game => {
      const gameId = game.url.split('/').pop();
      if (existingGameIds.has(gameId)) return false;
      if (lastKnownEndTime && game.end_time && game.end_time <= lastKnownEndTime) return false;
      return true;
    });
    
    if (newGames.length === 0) {
      ss.toast('No new games found (all were duplicates)!', 'ℹ️', 3);
      return;
    }
    
    ss.toast(`Found ${newGames.length} new games. Processing...`, '⏳', -1);
    const rows = processGamesData(newGames, username);
    
    if (rows.length > 0) {
      const lastRow = gamesSheet.getLastRow();
      gamesSheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
      
      // Update last known game URL to the newest by end_time among all new and existing
      let mostRecentGame = newGames[0];
      for (const game of newGames) {
        if (game.end_time > mostRecentGame.end_time) {
          mostRecentGame = game;
        }
      }
      props.setProperty('LAST_GAME_URL', mostRecentGame.url);
      
      ss.toast(`✅ Added ${rows.length} new games!`, '✅', 5);
      
      // Process auto-analysis and callback data for new games
      const gamesToProcess = newGames.map(g => {
        const gameId = g.url.split('/').pop();
        return {
          row: findGameRow(gameId),
          gameId: gameId,
          gameUrl: g.url,
          white: g.white?.username || '',
          black: g.black?.username || '',
          timeClass: getTimeClass(g.time_class),
          outcome: getGameOutcome(g, CONFIG.USERNAME),
          pgn: g.pgn || ''
        };
      }).filter(g => g.row > 0 && g.gameId && g.white && g.black);
      
      processNewGamesAutoFeatures(gamesToProcess);
      
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Error: ${error.message}`);
    Logger.log(error);
  }
}

// Fetch with ETag support
function fetchWithETag(url, etag) {
  const options = {
    muteHttpExceptions: true,
    headers: {}
  };
  
  if (etag) {
    options.headers['If-None-Match'] = etag;
  }
  
  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  
  if (code === 304) {
    // Not modified
    return { data: null, etag: etag };
  }
  
  if (code === 200) {
    const newETag = response.getHeaders()['ETag'] || response.getHeaders()['etag'] || '';
    const data = JSON.parse(response.getContentText());
    return { data: data, etag: newETag };
  }
  
  throw new Error(`HTTP ${code}: ${response.getContentText()}`);
}

// Convert archive URL to storage key
function archiveUrlToKey(url) {
  // Extract YYYY/MM from URL like https://api.chess.com/pub/player/username/games/2024/09
  const match = url.match(/(\d{4})\/(\d{2})$/);
  return match ? `${match[1]}_${match[2]}` : url;
}

// Parse time control string into components
function parseTimeControl(timeControl, timeClass) {
  const result = {
    type: timeClass === 'daily' ? 'Daily' : 'Live',
    baseTime: null,
    increment: null,
    correspondenceTime: null
  };
  
  if (!timeControl) return result;
  
  const tcStr = String(timeControl);
  
  // Check if correspondence/daily format (1/value)
  if (tcStr.includes('/')) {
    const parts = tcStr.split('/');
    if (parts.length === 2) {
      result.correspondenceTime = parseInt(parts[1]) || null;
    }
  }
  // Check if live format with increment (value+value)
  else if (tcStr.includes('+')) {
    const parts = tcStr.split('+');
    if (parts.length === 2) {
      result.baseTime = parseInt(parts[0]) || null;
      result.increment = parseInt(parts[1]) || null;
    }
  }
  // Simple live format (just value)
  else {
    result.baseTime = parseInt(tcStr) || null;
    result.increment = 0;
  }
  
  return result;
}

// Helper function to process games data
function processGamesData(games, username) {
  const rows = [];
  const derivedRows = [];
  
  // Sort games by timestamp (oldest first) to ensure Last Rating fills correctly
  const sortedGames = games.slice().sort((a, b) => a.end_time - b.end_time);
  
  // Pre-load existing games data once for performance
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const derivedSheet = ss.getSheetByName(SHEETS.DERIVED);
  let existingGames = [];
  
  if (gamesSheet && gamesSheet.getLastRow() > 1) {
    const data = gamesSheet.getDataRange().getValues();
    // Build lookup map: format -> array of {timestamp, rating}
    for (let i = 1; i < data.length; i++) {
      try {
        const format = data[i][7]; // Format column (index 7)
        const endDate = data[i][1]; // End Date (Date object)
        const endTime = data[i][2]; // End Time (Date object)
        const myRating = data[i][8]; // My Rating column (index 8)
        let timestamp = 0;
        if (endDate instanceof Date && endTime instanceof Date) {
          timestamp = new Date(
            endDate.getFullYear(), endDate.getMonth(), endDate.getDate(),
            endTime.getHours(), endTime.getMinutes(), endTime.getSeconds()
          ).getTime() / 1000;
        } else if (endDate instanceof Date) {
          timestamp = Math.floor(endDate.getTime() / 1000);
        }
        
        existingGames.push({
          format: format,
          timestamp: timestamp,
          rating: myRating
        });
      } catch (error) {
        // Skip malformed rows
        continue;
      }
    }
  }
  
  for (const game of sortedGames) {
    try {
      if (!game || !game.url || !game.end_time) {
        Logger.log('Skipping game with missing data');
        continue;
      }
      
      const endDate = new Date(game.end_time * 1000);
      const startDate = extractStartFromPGN(game.pgn) || null;
      const gameId = game.url.split('/').pop();
      const eco = extractECOFromPGN(game.pgn);
      const ecoUrl = extractECOUrlFromPGN(game.pgn);
      const ecoSlug = extractECOSlug(ecoUrl);
      const outcome = getGameOutcome(game, username);
      const termination = getGameTermination(game, username);
      const format = getGameFormat(game);
      const timeClass = getTimeClass(game.time_class);
      const duration = extractDurationFromPGN(game.pgn);
      
      // Determine my color and opponent
      const isWhite = game.white?.username.toLowerCase() === username.toLowerCase();
      const myColor = isWhite ? 'white' : 'black';
      const opponent = isWhite ? game.black?.username : game.white?.username;
      const myRating = isWhite ? game.white?.rating : game.black?.rating;
      const oppRating = isWhite ? game.black?.rating : game.white?.rating;
      
      // Calculate Last Rating from pre-loaded data AND games processed in this batch
      // Use end_time comparison; fall back to earlier same-second using <=
      let lastRating = null;
      let lastGameTime = 0;
      for (const existingGame of existingGames) {
        if (existingGame.format === format) {
          const isEarlier = existingGame.timestamp < game.end_time || (existingGame.timestamp === game.end_time);
          if (isEarlier && existingGame.timestamp >= lastGameTime && existingGame.rating && existingGame.rating !== 'N/A') {
            lastGameTime = existingGame.timestamp;
            lastRating = existingGame.rating;
          }
        }
      }
      
      // Parse time control
      const tcParsed = parseTimeControl(game.time_control, game.time_class);
      
      // Extract compact move data
      const tcn = game.tcn || extractTCNFromPGN(game.pgn) || '';
      const moveData = extractMovesWithClocks(game.pgn, tcParsed.baseTime, tcParsed.increment);
      const mc36 = encodeClocksBase36(moveData.clocks);
      
      // Combined Start/End datetimes
      const endDateTime = new Date(endDate.getTime());
      const startDateTime = startDate ? new Date(startDate.getTime()) : (duration && duration > 0 ? new Date(endDateTime.getTime() - (duration * 1000)) : null);
      
      // Fill legacy Games sheet row to preserve current flows
      const endDateObj = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
      const endTimeObj = new Date(1970, 0, 1, endDate.getHours(), endDate.getMinutes(), endDate.getSeconds());
      rows.push([
        game.url, endDateObj, endTimeObj, myColor, opponent || 'Unknown',
        outcome, termination, format,
        myRating || 'N/A', oppRating || 'N/A', lastRating || 'N/A',
        gameId, false, false
      ]);
      
      // Calculate Moves (ply count / 2 rounded up)
      const movesCount = moveData.plyCount > 0 ? Math.ceil(moveData.plyCount / 2) : 0;
      
      // Store combined lean data in derived sheet
      const dbValues = getOpeningOutputs(ecoSlug);
      derivedRows.push([
        gameId,
        game.url,
        startDateTime,
        endDateTime,
        duration || null,
        timeClass.toLowerCase() !== 'daily',
        timeClass,
        tcParsed.baseTime,
        tcParsed.increment,
        isWhite,
        opponent || 'Unknown',
        myRating || 'N/A',
        oppRating || 'N/A',
        outcome,
        termination,
        eco,
        dbValues[0], // Opening Name
        dbValues[1], // Opening Family
        moveData.plyCount,
        movesCount,
        tcn,
        mc36
      ]);
      
      // Add this game to existingGames for subsequent games in this batch
      existingGames.push({
        format: format,
        timestamp: game.end_time,
        rating: myRating
      });
      
    } catch (error) {
      Logger.log(`Error processing game ${game?.url}: ${error.message}`);
      continue;
    }
  }
  
  // Write derived data to hidden sheet
  if (derivedSheet && derivedRows.length > 0) {
    // Ensure DB columns exist before writing, to match value count
    ensureDerivedDbColumns(derivedSheet);
    const lastRow = derivedSheet.getLastRow();
    derivedSheet.getRange(lastRow + 1, 1, derivedRows.length, derivedRows[0].length).setValues(derivedRows);
  }
  
  return rows;
}

function ensureDerivedDbColumns(derivedSheet) {
  if (!derivedSheet) return;
  const lastCol = derivedSheet.getLastColumn();
  const headers = derivedSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (const header of DERIVED_DB_HEADERS) {
    const idx = headers.indexOf(header);
    if (idx === -1) {
      derivedSheet.insertColumnAfter(derivedSheet.getLastColumn());
      const col = derivedSheet.getLastColumn();
      derivedSheet.getRange(1, col).setValue(header)
        .setFontWeight('bold')
        .setBackground('#666666')
        .setFontColor('#ffffff');
    }
  }
}

// Get game format based on rules and time control
function getGameFormat(game) {
  const rules = game.rules || 'chess';
  const timeClass = game.time_class || '';
  
  if (rules === 'chess') {
    // Use time class for standard chess (Bullet, Blitz, Rapid, Daily)
    return getTimeClass(timeClass);
  } else if (rules === 'chess960') {
    return timeClass === 'daily' ? 'daily960' : 'live960';
  } else {
    // For other variants, return the rules name
    return rules;
  }
}

// Remove duplicate games based on Game ID
function removeDuplicates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  
  if (!gamesSheet) {
    SpreadsheetApp.getUi().alert('❌ Games sheet not found!');
    return;
  }
  
  const data = gamesSheet.getDataRange().getValues();
  const header = data[0];
  const gameIdCol = 11; // Game ID column (index 11)
  
  const seen = new Set();
  const rowsToKeep = [header];
  let duplicateCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const gameId = data[i][gameIdCol];
    
    if (!seen.has(gameId)) {
      seen.add(gameId);
      rowsToKeep.push(data[i]);
    } else {
      duplicateCount++;
    }
  }
  
  if (duplicateCount > 0) {
    gamesSheet.clear();
    gamesSheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
    
    gamesSheet.getRange(1, 1, 1, header.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    gamesSheet.setFrozenRows(1);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Removed ${duplicateCount} duplicate(s)`, 
      '🗑️', 
      3
    );
    Logger.log(`Removed ${duplicateCount} duplicates`);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('No duplicates found!', 'ℹ️', 2);
  }
}

// ============================================
// HELPER FUNCTIONS
// ============================================

function extractMovesFromPGN(pgn) {
  const moveSection = pgn.split(/\n\n/)[1] || pgn;
  
  return moveSection
    .replace(/\{[^}]*\}/g, '')
    .replace(/\([^)]*\)/g, '')
    .replace(/\[[^\]]*\]/g, '')
    .replace(/\$\d+/g, '')
    .replace(/\d+\.{3}/g, '')
    .replace(/\d+\./g, '')
    .replace(/[!?+#]+/g, '')
    .trim()
    .split(/\s+/)
    .filter(m => m && m !== '*' && !m.match(/^(1-0|0-1|1\/2-1\/2)$/));
}

function extractECOFromPGN(pgn) {
  if (!pgn) return '';
  const match = pgn.match(/\[ECO "([^"]+)"\]/);
  return match ? match[1] : '';
}

// Extract ECO URL from PGN
function extractECOUrlFromPGN(pgn) {
  if (!pgn) return '';
  const match = pgn.match(/\[ECOUrl "([^"]+)"\]/);
  return match ? match[1] : '';
}

/**
 * Extract opening slug from ECO URL
 * Rules:
 * 1. Remove base URL (keep only the slug part)
 * 2. Remove move sequences (patterns like -3...Nf6-4.g3)
 * 3. Keep "with-X-move" patterns (e.g., with-7-a5)
 * 4. Keep "with-X-move-and-Y-move" patterns (e.g., with-2-d4-and-3-g3)
 * 5. Remove anything after move numbers that aren't part of "with" patterns
 */
function extractECOSlug(ecoUrl) {
  if (!ecoUrl || !ecoUrl.includes('chess.com/openings/')) return '';
  
  // Extract the slug part after '/openings/'
  const slug = ecoUrl.split('/openings/')[1] || '';
  if (!slug) return '';
  
  // Strategy: Find the first move sequence that's NOT part of a "with" pattern and trim from there
  
  // Pattern 1: with-NUMBER-MOVE-and-NUMBER-MOVE (keep this entire pattern)
  // Pattern 2: with-NUMBER-MOVE (keep this entire pattern)
  // Pattern 3: -NUMBER (where NUMBER is followed by . or ... indicating moves) - REMOVE from here onward
  
  // First, protect "with" patterns by replacing them temporarily
  let protected = slug;
  const withPatterns = [];
  
  // Match: with-NUMBER-MOVE-and-NUMBER-MOVE
  // Move can be: standard notation (Nf3, e4, etc.) or castling (O-O, O-O-O)
  const withAndPattern = /with-(\d+)-(O-O(?:-O)?|[a-zA-Z0-9]+)-and-(\d+)-(O-O(?:-O)?|[a-zA-Z0-9]+)/g;
  protected = protected.replace(withAndPattern, (match) => {
    const placeholder = `__WITH_AND_${withPatterns.length}__`;
    withPatterns.push(match);
    return placeholder;
  });
  
  // Match: with-NUMBER-MOVE (but not followed by -and-)
  // Move can be: standard notation (Nf3, e4, etc.) or castling (O-O, O-O-O)
  const withPattern = /with-(\d+)-(O-O(?:-O)?|[a-zA-Z0-9]+)(?!-and-)/g;
  protected = protected.replace(withPattern, (match) => {
    const placeholder = `__WITH_${withPatterns.length}__`;
    withPatterns.push(match);
    return placeholder;
  });
  
  // Now find the first move sequence indicator
  // Look for patterns like: -3...Nf6 or -4.g3 or -7...g6 or ...8.Nf3 or ...5.cxd4 or ...e6
  // These indicate the start of move notation
  // Pattern matches: 
  //   -NUMBER. or -NUMBER... (dash followed by move number)
  //   ...NUMBER. (three dots followed by move number)
  //   ...[a-zA-Z] (three dots followed by move notation without number)
  const movePattern = /(-\d+\.{0,3}[a-zA-Z]|\.{3}\d+\.|\.{3}[a-zA-Z])/;
  const moveMatch = protected.match(movePattern);
  
  if (moveMatch) {
    // Trim from the first move sequence
    protected = protected.substring(0, moveMatch.index);
  }
  
  // Restore "with" patterns
  for (let i = 0; i < withPatterns.length; i++) {
    protected = protected.replace(`__WITH_AND_${i}__`, withPatterns[i]);
    protected = protected.replace(`__WITH_${i}__`, withPatterns[i]);
  }
  
  return protected;
}

// ================================
// OPENINGS DB LOOKUP (by ECO Slug)
// ================================

function loadOpeningsDbCache() {
  if (OPENINGS_DB_CACHE) return OPENINGS_DB_CACHE;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName(SHEETS.OPENINGS_DB);
  const cache = new Map();
  if (!dbSheet) {
    OPENINGS_DB_CACHE = cache;
    return cache;
  }
  const values = dbSheet.getDataRange().getValues();
  if (!values || values.length < 2) {
    OPENINGS_DB_CACHE = cache;
    return cache;
  }
  const header = values[0];
  const slugIdx = header.indexOf('Trim Slug');
  const familyIdx = header.indexOf('Family');
  // The TSV has two 'Name' headers. We'll treat the first as Full Name and the second as Base Name.
  // Positions per OPENINGS_DB_HEADERS:
  // 0: Name (Full), 1: Trim Slug, 2: Family, 3: Name (Base), 4..9: Variation 1..6
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const trimSlug = String(row[1] || '').trim();
    if (!trimSlug) continue;
    const fullName = String(row[0] || '');
    const baseName = String(row[3] || '');
    const family = String(row[2] || '');
    const v1 = String(row[4] || '');
    const v2 = String(row[5] || '');
    const v3 = String(row[6] || '');
    const v4 = String(row[7] || '');
    const v5 = String(row[8] || '');
    const v6 = String(row[9] || '');
    cache.set(trimSlug, [fullName, family, baseName, v1, v2, v3, v4, v5, v6]);
  }
  OPENINGS_DB_CACHE = cache;
  return cache;
}

function normalizeSlugForDb(ecoSlug) {
  if (!ecoSlug) return '';
  // The DB uses Title-Case with hyphens; ecoSlug looks similar but may include lowercase and numbers like with-3-Nc3
  // We'll convert to Title-Case tokens separated by '-' and ensure castling tokens are normalized.
  const tokens = ecoSlug
    .replace(/_/g, '-')
    .split('-')
    .filter(Boolean)
    .map(tok => {
      if (/^with$/i.test(tok) || /^and$/i.test(tok)) return tok.charAt(0).toUpperCase() + tok.slice(1).toLowerCase();
      if (/^o$/i.test(tok)) return 'O';
      if (/^o\so$/i.test(tok)) return 'O-O';
      // Preserve chess move tokens/case like Nf3, e4, O-O, but capitalize words
      if (/^[a-z][a-z]+$/i.test(tok)) {
        return tok.charAt(0).toUpperCase() + tok.slice(1);
      }
      return tok;
    });
  return tokens.join('-');
}

function getDbMappingValues(ecoSlug) {
  // Returns array matching DERIVED_DB_HEADERS order
  const empty = ['', '', '', '', '', '', '', '', ''];
  if (!ecoSlug) return empty;
  const db = loadOpeningsDbCache();
  // Try direct match first
  if (db.has(ecoSlug)) return db.get(ecoSlug);
  // Try normalized form
  const normalized = normalizeSlugForDb(ecoSlug);
  if (db.has(normalized)) return db.get(normalized);
  // Try loosening: drop trailing move qualifiers like 'with-3-Nc3' if not found
  const withoutWith = ecoSlug.split('-with-')[0];
  if (withoutWith && db.has(withoutWith)) return db.get(withoutWith);
  const normalizedWithoutWith = normalized.split('-with-')[0];
  if (normalizedWithoutWith && db.has(normalizedWithoutWith)) return db.get(normalizedWithoutWith);
  return empty;
}

// Minimal opening outputs: [Opening Name, Opening Family]
function getOpeningOutputs(ecoSlug) {
  const db = loadOpeningsDbCache();
  if (!ecoSlug) return ['', ''];
  const keys = [ecoSlug, normalizeSlugForDb(ecoSlug), ecoSlug.split('-with-')[0] || '', normalizeSlugForDb(ecoSlug).split('-with-')[0] || ''];
  for (const k of keys) {
    if (k && db.has(k)) {
      const row = db.get(k);
      return [row[0] || '', row[2] || '']; // Full Name, Family
    }
  }
  return ['', ''];
}

// Extract start datetime from PGN headers if available
function extractStartFromPGN(pgn) {
  if (!pgn) return null;
  const dateMatch = pgn.match(/\[UTCDate "([^"]+)"\]/);
  const timeMatch = pgn.match(/\[UTCTime "([^"]+)"\]/);
  if (!dateMatch || !timeMatch) return null;
  try {
    const d = dateMatch[1].split('.');
    const t = timeMatch[1].split(':');
    return new Date(Date.UTC(parseInt(d[0]), parseInt(d[1]) - 1, parseInt(d[2]), parseInt(t[0]), parseInt(t[1]), parseInt(t[2])));
  } catch (e) {
    return null;
  }
}

// Try to extract a TCN-like move list from PGN (fallback)
function extractTCNFromPGN(_pgn) { return ''; }

// Encode clocks (seconds) to base36 deciseconds dot-joined
function encodeClocksBase36(clocksCsv) {
  if (!clocksCsv) return '';
  const parts = String(clocksCsv).split(',').map(s => s.trim()).filter(Boolean);
  if (parts.length === 0) return '';
  return parts.map(p => {
    const ds = Math.round(parseFloat(p) * 10);
    const val = isFinite(ds) && ds >= 0 ? ds : 0;
    return val.toString(36);
  }).join('.');
}

// Extract moves with clock times from PGN
function extractMovesWithClocks(pgn, baseTime, increment) {
  if (!pgn) return { moves: [], clocks: [], times: [] };
  
  const moveSection = pgn.split(/\n\n/)[1] || pgn;
  const moves = [];
  const clocks = [];
  const times = [];
  
  // Regex to match move and its clock: "e4 {[%clk 0:02:59.9]}"
  const movePattern = /([NBRQK]?[a-h]?[1-8]?x?[a-h][1-8](?:=[NBRQK])?|O-O(?:-O)?)\s*\{?\[%clk\s+(\d+):(\d+):(\d+)(?:\.(\d+))?\]?\}?/g;
  
  let match;
  let prevClock = [baseTime || 0, baseTime || 0]; // [white, black] previous clocks
  let moveIndex = 0;
  
  while ((match = movePattern.exec(moveSection)) !== null) {
    const move = match[1];
    const hours = parseInt(match[2]) || 0;
    const minutes = parseInt(match[3]) || 0;
    const seconds = parseInt(match[4]) || 0;
    const deciseconds = parseInt(match[5]) || 0;
    
    // Convert clock to total seconds
    const clockSeconds = hours * 3600 + minutes * 60 + seconds + deciseconds / 10;
    
    moves.push(move);
    clocks.push(clockSeconds);
    
    // Calculate time spent on this move
    const playerIndex = moveIndex % 2; // 0 = white, 1 = black
    const prevPlayerClock = prevClock[playerIndex];
    
    // Time spent = previous clock - current clock + increment
    let timeSpent = prevPlayerClock - clockSeconds + (increment || 0);
    // Allow 0.0 seconds moves (e.g., premove)
    if (timeSpent < 0) timeSpent = 0;
    
    times.push(Math.round(timeSpent * 10) / 10); // Round to 1 decimal
    
    // Update previous clock for this player
    prevClock[playerIndex] = clockSeconds;
    
    moveIndex++;
  }
  
  return { 
    moveList: moves.join(', '), 
    clocks: clocks.join(', '), 
    times: times.join(', '),
    plyCount: moves.length
  };
}

function extractDurationFromPGN(pgn) {
  if (!pgn) return null;
  
  const dateMatch = pgn.match(/\[UTCDate "([^"]+)"\]/);
  const timeMatch = pgn.match(/\[UTCTime "([^"]+)"\]/);
  const endDateMatch = pgn.match(/\[EndDate "([^"]+)"\]/);
  const endTimeMatch = pgn.match(/\[EndTime "([^"]+)"\]/);
  
  if (!dateMatch || !timeMatch || !endDateMatch || !endTimeMatch) {
    return null;
  }
  
  try {
    const startDateParts = dateMatch[1].split('.');
    const startTimeParts = timeMatch[1].split(':');
    const startDate = new Date(Date.UTC(
      parseInt(startDateParts[0]),
      parseInt(startDateParts[1]) - 1,
      parseInt(startDateParts[2]),
      parseInt(startTimeParts[0]),
      parseInt(startTimeParts[1]),
      parseInt(startTimeParts[2])
    ));
    
    const endDateParts = endDateMatch[1].split('.');
    const endTimeParts = endTimeMatch[1].split(':');
    const endDate = new Date(Date.UTC(
      parseInt(endDateParts[0]),
      parseInt(endDateParts[1]) - 1,
      parseInt(endDateParts[2]),
      parseInt(endTimeParts[0]),
      parseInt(endTimeParts[1]),
      parseInt(endTimeParts[2])
    ));
    
    const durationMs = endDate.getTime() - startDate.getTime();
    return Math.round(durationMs / 1000);
  } catch (error) {
    Logger.log(`Error parsing duration: ${error.message}`);
    return null;
  }
}

function getGameOutcome(game, username) {
  if (!game || !game.white || !game.black) return 'Unknown';
  
  const isWhite = game.white.username?.toLowerCase() === username.toLowerCase();
  const myResult = isWhite ? game.white.result : game.black.result;
  
  if (!myResult) return 'Unknown';
  
  return RESULT_MAP[myResult] || 'Unknown';
}

function getGameTermination(game, username) {
  if (!game || !game.white || !game.black) return 'Unknown';
  
  const isWhite = game.white.username?.toLowerCase() === username.toLowerCase();
  const myResult = isWhite ? game.white.result : game.black.result;
  const opponentResult = isWhite ? game.black.result : game.white.result;
  
  if (!myResult) return 'Unknown';
  
  // If I won, use opponent's result for termination
  if (myResult === 'win') {
    return TERMINATION_MAP[opponentResult] || opponentResult;
  }
  
  // Otherwise use my result
  return TERMINATION_MAP[myResult] || myResult;
}

function getTimeClass(timeClass) {
  if (timeClass === 'bullet') return 'Bullet';
  if (timeClass === 'blitz') return 'Blitz';
  if (timeClass === 'rapid') return 'Rapid';
  if (timeClass === 'daily') return 'Daily';
  return timeClass || 'Unknown';
}
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  if (!gamesSheet) {
    gamesSheet = ss.insertSheet(SHEETS.GAMES);
    const headers = [
      'Game URL', 'End Date', 'End Time', 'My Color', 'Opponent',
      'Outcome', 'Termination', 'Format',
      'My Rating', 'Opp Rating', 'Last Rating',
      'Game ID', 'Analyzed', 'Callback Fetched'
    ];
    gamesSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    gamesSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    gamesSheet.setFrozenRows(1);
    gamesSheet.setColumnWidth(1, 200);
    
    // Format date and time columns
    gamesSheet.getRange('B:B').setNumberFormat('m"/"d"/"yy');
    gamesSheet.getRange('C:C').setNumberFormat('h:mm AM/PM');
  }
  
  let derivedSheet = ss.getSheetByName(SHEETS.DERIVED);
  if (!derivedSheet) {
    derivedSheet = ss.insertSheet(SHEETS.DERIVED);
    const headers = [
      // Combined lean schema
      'Game ID', 'Game URL',
      'Start', 'End', 'Duration (s)',
      'Is Live', 'Time Class', 'Base Time (s)', 'Increment (s)',
      'Is White', 'Opponent', 'My Rating', 'Opp Rating',
      'Outcome', 'Termination',
      'ECO', 'Opening Name', 'Opening Family',
      'Ply Count', 'Moves',
      // Compact detail
      'tcn', 'mc36'
    ];
    derivedSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    derivedSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#666666')
      .setFontColor('#ffffff');
    derivedSheet.setFrozenRows(1);
    
    // Format datetime columns
    const startCol = 3;
    const endCol = 4;
    derivedSheet.getRange(1, startCol, derivedSheet.getMaxRows(), 1).setNumberFormat('m"/"d"/"yy h:mm AM/PM');
    derivedSheet.getRange(1, endCol, derivedSheet.getMaxRows(), 1).setNumberFormat('m"/"d"/"yy h:mm AM/PM');
    
    // Hide the derived sheet
    derivedSheet.hideSheet();
  } else {
    // If a legacy sheet exists, we won't auto-migrate columns here.
  }

  
  // Apply formatting to existing Games sheet if it exists
  if (gamesSheet) {
    gamesSheet.getRange('B:B').setNumberFormat('m"/"d"/"yy');
    gamesSheet.getRange('C:C').setNumberFormat('h:mm AM/PM');
  }
  
   let callbackSheet = ss.getSheetByName(SHEETS.CALLBACK);
  if (!callbackSheet) {
    callbackSheet = ss.insertSheet(SHEETS.CALLBACK);
    const headers = [
      'Game ID', 'Game URL', 'Callback URL', 'End Time', 'My Color', 'Time Class',
      'My Rating', 'Opp Rating', 'My Rating Change', 'Opp Rating Change',
      'My Rating Before', 'Opp Rating Before',
      'Base Time', 'Time Increment', 'Move Timestamps',
      'My Username', 'My Country', 'My Membership', 'My Member Since',
      'My Default Tab', 'My Post Move Action', 'My Location',
      'Opp Username', 'Opp Country', 'Opp Membership', 'Opp Member Since',
      'Opp Default Tab', 'Opp Post Move Action', 'Opp Location',
      'Date Fetched'
    ];
    callbackSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    callbackSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#f4b400')
      .setFontColor('#ffffff');
    callbackSheet.setFrozenRows(1);
    
    callbackSheet.getRange('K:K').setNumberFormat('@STRING@');
  }

  // Ensure Openings DB sheet exists with headers for paste/import
  let dbSheet = ss.getSheetByName(SHEETS.OPENINGS_DB);
  if (!dbSheet) {
    dbSheet = ss.insertSheet(SHEETS.OPENINGS_DB);
    dbSheet.getRange(1, 1, 1, OPENINGS_DB_HEADERS.length).setValues([OPENINGS_DB_HEADERS]);
    dbSheet.getRange(1, 1, 1, OPENINGS_DB_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#0b8043')
      .setFontColor('#ffffff');
    dbSheet.setFrozenRows(1);
  }
  
  SpreadsheetApp.getUi().alert('✅ Sheets setup complete!');
}




// ============================================
// UTILITIES: Refresh mappings across existing rows
// ============================================
function refreshDerivedDbMappings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const derivedSheet = ss.getSheetByName(SHEETS.DERIVED);
  if (!derivedSheet) {
    SpreadsheetApp.getUi().alert('Derived sheet not found');
    return;
  }
  // Ensure cache is loaded fresh
  OPENINGS_DB_CACHE = null;
  loadOpeningsDbCache();

  const lastRow = derivedSheet.getLastRow();
  const lastCol = derivedSheet.getLastColumn();
  if (lastRow < 2) return;

  const headers = derivedSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const ecoSlugCol = headers.indexOf('ECO Slug') + 1;
  if (ecoSlugCol <= 0) {
    SpreadsheetApp.getUi().alert('ECO Slug column not found');
    return;
  }

  // Determine starting column for DB headers, ensuring they exist
  let startDbCol = -1;
  for (let i = 0; i < DERIVED_DB_HEADERS.length; i++) {
    const idx = headers.indexOf(DERIVED_DB_HEADERS[i]);
    if (idx >= 0) {
      startDbCol = startDbCol === -1 ? (idx + 1) : Math.min(startDbCol, idx + 1);
    }
  }
  if (startDbCol === -1) {
    // Append columns if missing
    let currentHeaders = headers.slice();
    for (const header of DERIVED_DB_HEADERS) {
      currentHeaders = derivedSheet.getRange(1, 1, 1, derivedSheet.getLastColumn()).getValues()[0];
      if (!currentHeaders.includes(header)) {
        derivedSheet.insertColumnAfter(derivedSheet.getLastColumn());
        const col = derivedSheet.getLastColumn();
        derivedSheet.getRange(1, col).setValue(header)
          .setFontWeight('bold')
          .setBackground('#666666')
          .setFontColor('#ffffff');
      }
    }
    // Recompute
    const newHeaders = derivedSheet.getRange(1, 1, 1, derivedSheet.getLastColumn()).getValues()[0];
    startDbCol = newHeaders.indexOf(DERIVED_DB_HEADERS[0]) + 1;
  }

  const ecoSlugs = derivedSheet.getRange(2, ecoSlugCol, lastRow - 1, 1).getValues().map(r => String(r[0] || ''));
  const writeRows = [];
  for (const ecoSlug of ecoSlugs) {
    const vals = getDbMappingValues(ecoSlug);
    writeRows.push(vals);
  }

  derivedSheet.getRange(2, startDbCol, writeRows.length, DERIVED_DB_HEADERS.length).setValues(writeRows);
  SpreadsheetApp.getUi().alert('✅ Opening mappings refreshed');
}
