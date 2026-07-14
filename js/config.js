const CONFIG = {
    // 請在此更換為您部署後的 Google Apps Script Web App URL
    API_URL: "https://script.google.com/macros/s/AKfycbylJdE0zrBJkgbjkvXKKCbhQYs2wnrHCTMutJueWPRISeiUcAuQNYBf-Mg5X9-1GEKZ0w/exec",

    // 目前預設日期 (用於篩選資料)
    DEFAULT_DATE: new Date().toLocaleDateString('sv'), // yyyy-mm-dd

    TEAMS: ["藍鳥隊", "黑鳥隊", "青鳥隊", "粉鳥隊"],
    AREAS: ["猛禽", "小鳥", "鳥蛋"],
    SHEET_POINTS: "積點紀錄",

    // 隊伍顏色映射 (用於 UI 呈現)
    TEAM_COLORS: {
        "藍鳥隊": "#4a90e2",
        "黑鳥隊": "#6c757d",
        "青鳥隊": "#20c997",
        "粉鳥隊": "#e91e63"
    },

    AREA_COLORS: {
        "猛禽": "#ff6b35",
        "小鳥": "#4ecdc4",
        "鳥蛋": "#ffe66d",
        "猛禽區": "#ff6b35",
        "小鳥區": "#4ecdc4",
        "鳥蛋區": "#ffe66d"
    },

    // 🏆 新增：賽事類型與其分組積分規則
    TOURNAMENT_MODES: {
        "default": { name: "預設模式", points: { "猛禽": 100, "小鳥": 80, "鳥蛋": 60 } },
        "egg": { name: "孵蛋模式 (猛蛋混)", points: { "猛禽+鳥蛋": 80, "小鳥": 80 } },
        "fox": { name: "狐狸模式 (猛鳥混)", points: { "猛禽+小鳥": 90, "鳥蛋": 60 } },
        "duck": { name: "醬板鴨模式 (小蛋混)", points: { "猛禽": 100, "小鳥+鳥蛋": 70 } }
    },

    // 🏆 新增：鳥樂特別賽規則
    LOTTERY_CONFIG: {
        小組賽: { win: 100, lose: 50 },
        淘汰賽: { 1: 300, 2: 250, 3: 200, 4: 150, 5: 100, 6: 100 }
    }
};
