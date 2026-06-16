const Viewer = {
    isBroadcastMode: false,

    getTeamRoster(match, side, separator = "/") {
        return [match[`${side}隊員1`], match[`${side}隊員2`], match[`${side}隊員3`]]
            .filter(name => name && name !== "待定")
            .join(separator);
    },

    normalizeMatchText(value) {
        return String(value || "").replace(/\s+/g, " ").trim();
    },

    getChasingGroupKey(match) {
        return [
            this.normalizeMatchText(match["比賽時間"]),
            this.normalizeMatchText(match["場地"]),
            this.normalizeMatchText(match["區"]),
            this.normalizeMatchText(match["A隊名"]),
            this.normalizeMatchText(match["B隊名"])
        ].join("||");
    },

    isRelayRound(match) {
        const roundText = String((match && match["輪次"]) || "");
        return roundText.includes("接力") || roundText.includes("追分");
    },

    groupChasingMatches(matches) {
        const groups = new Map();

        (matches || []).forEach(match => {
            const key = this.isRelayRound(match)
                ? this.getChasingGroupKey(match)
                : [
                    this.getChasingGroupKey(match),
                    this.normalizeMatchText(match["輪次"]),
                    this.normalizeMatchText(match["序號"])
                ].join("||");
            if (!groups.has(key)) groups.set(key, []);
            groups.get(key).push(match);
        });

        return [...groups.values()]
            .map(group => group.sort((a, b) => {
                const seqA = parseInt(a["序號"], 10) || 0;
                const seqB = parseInt(b["序號"], 10) || 0;
                return seqA - seqB;
            }))
            .sort((groupA, groupB) => {
                const firstA = groupA[0] || {};
                const firstB = groupB[0] || {};
                const timeDiff = this.normalizeMatchText(firstA["比賽時間"]).localeCompare(this.normalizeMatchText(firstB["比賽時間"]), "zh-Hant");
                if (timeDiff !== 0) return timeDiff;

                const seqA = parseInt(firstA["序號"], 10) || 0;
                const seqB = parseInt(firstB["序號"], 10) || 0;
                if (seqA !== seqB) return seqA - seqB;

                return this.normalizeMatchText(firstA["區"]).localeCompare(this.normalizeMatchText(firstB["區"]), "zh-Hant");
            });
    },

    getChasingSequenceLabel(matches, prefix = "序號: ") {
        const seqList = [...new Set((matches || [])
            .map(match => String(match["序號"] || "").trim())
            .filter(Boolean))]
            .sort((a, b) => (parseInt(a, 10) || 0) - (parseInt(b, 10) || 0));

        if (seqList.length === 0) return "追分/淘汰";
        return `${prefix}${seqList.join(", ")}`;
    },

    isChasingSeries(matches) {
        return (matches || []).some(match => this.isRelayRound(match));
    },

    toggleBroadcastMode(checked) {
        this.isBroadcastMode = checked;
        // 同步另一個切換按鈕 (如果存在)
        const otherToggle = document.getElementById("broadcast-mode-ch-toggle");
        if (otherToggle) otherToggle.checked = checked;
        
        this.loadSchedule(true);
    },
    async loadSchedule(silent = false) {
        const rrContainer = document.getElementById("v-rr-list-container");
        const chasingContainer = document.getElementById("v-chasing-list-container");
        
        if (!rrContainer || !chasingContainer) return;

        if (!silent) {
            rrContainer.innerHTML = "<div style='text-align:center; padding: 2rem; width: 100%;'><i class='fas fa-spinner fa-spin fa-2x'></i><br>載入中...</div>";
            chasingContainer.innerHTML = "<div style='text-align:center; padding: 2rem; width: 100%;'><i class='fas fa-spinner fa-spin fa-2x'></i><br>載入中...</div>";
        }

        try {
            const rrData = await API.getSchedule();
            const chData = await API.getChasingSchedule();
            
            // 渲染循環賽 (Cards)
            if (rrData && rrData.data) {
                rrContainer.innerHTML = "";
                if (rrData.data.length === 0) {
                    rrContainer.innerHTML = "<div class='card' style='text-align:center; width:100%;'>目前尚無賽程資料</div>";
                } else {
                    rrData.data.forEach(m => {
                        const status = m["比賽狀態"] || "待賽";
                        const isDone = status.includes("完賽");
                        const isLive = status.includes("進行中");
                        const statusHtml = isDone ? `<span class="status-badge status-done">已完賽</span>` : (isLive ? `<span class="status-badge status-live" style="background:#ff4757; color:white; animation: pulse 1.5s infinite;">即時比分</span>` : `<span class="status-badge status-pending">${status}</span>`);
                        
                        // User requested to show scores even if not finished
                        const aScore = m["A隊比分"] || 0;
                        const bScore = m["B隊比分"] || 0;
                        
                        rrContainer.innerHTML += `
                            <div class="card match-card" style="padding: 1.2rem; border-left: 4px solid var(--primary);">
                                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom: 0.8rem; padding-bottom: 0.5rem; border-bottom: 1px solid rgba(255,255,255,0.1);">
                                    <span style="color:var(--text-dim); font-size: 0.9rem;"><i class="far fa-clock"></i> ${m["比賽時間"] || ""}</span>
                                    <span style="color:var(--primary); font-weight:bold;">序號: ${m["序號"] || ""}</span>
                                    ${statusHtml}
                                </div>
                                <div style="color: var(--raptor); font-size: 0.85rem; margin-bottom:0.5rem;">
                                    第 ${m["輪次"]} 輪 - ${m["區"]} (${m["場地"]}場)
                                </div>
                                <div style="display:flex; justify-content:space-between; align-items:center; gap: 10px;">
                                    <div style="text-align:center; flex:1;">
                                        <div style="font-weight:bold; font-size:1.1rem; color:white;">${m["A隊名"] || ""}</div>
                                        <small style="color:var(--text-dim);">${m["A隊員1"] || ""}, ${m["A隊員2"] || ""}</small>
                                    </div>
                                    <div style="font-size:1.8rem; font-weight:bold; color:var(--accent); min-width: 80px; text-align:center; padding: 5px 10px; background:rgba(0,0,0,0.3); border-radius:8px;">
                                        ${aScore} : ${bScore}
                                    </div>
                                    <div style="text-align:center; flex:1;">
                                        <div style="font-weight:bold; font-size:1.1rem; color:white;">${m["B隊名"] || ""}</div>
                                        <small style="color:var(--text-dim);">${m["B隊員1"] || ""}, ${m["B隊員2"] || ""}</small>
                                    </div>
                                </div>
                                ${m["裁判"] ? `<div style="text-align:right; margin-top:0.8rem; font-size:0.85rem; color:var(--text-dim); border-top: 1px dashed rgba(255,255,255,0.05); padding-top:0.5rem;">裁判: ${m["裁判"]}</div>` : ''}
                            </div>
                        `;
                    });
                }
            }

            // 渲染追分賽與冠軍戰 (表格化顯示取分)
            if (chData && chData.data) {
                chasingContainer.innerHTML = "";
                if (chData.data.length === 0) {
                    chasingContainer.innerHTML = "<div class='card' style='text-align:center; width:100%;'>目前尚無追分/淘汰賽資料</div>";
                } else {
                    const groups = this.groupChasingMatches(chData.data);

                    groups.forEach(matches => {
                        const first = matches[0];
                        const last = matches[matches.length - 1];
                        
                        const status = last["比賽狀態"] || "待賽";
                        const isDone = status.includes("完賽");
                        const isLive = status.includes("進行中");
                        const statusHtml = isDone ? `<span class="status-badge status-done">已完賽</span>` : (isLive ? `<span class="status-badge status-live" style="background:#ff4757; color:white; animation: pulse 1.5s infinite;">即時比分</span>` : `<span class="status-badge status-pending">${status}</span>`);

                        const aScore = last["A隊比分"] || 0;
                        const bScore = last["B隊比分"] || 0;

                        // 判斷區塊顏色
                        const areaColor = first["區"] && first["區"].includes("猛禽") ? "var(--raptor)" : (first["區"] && first["區"].includes("小鳥") ? "var(--birdie)" : "var(--accent)");
                        const borderStyle = `border-left: 4px solid ${areaColor};`;

                        // 產生接力明細 (Legs)
                        let legInfoHtml = "";
                        let prevA = 0;
                        let prevB = 0;
                        
                        const isRelay = this.isChasingSeries(matches);

                        // 產生 Leg 資訊 (通用邏輯)
                        matches.forEach(m => {
                            const mStatus = m["比賽狀態"] || "待賽";
                            const mDone = mStatus.includes("完賽");
                            const mLive = mStatus.includes("進行中");
                            
                            const currA = parseInt(m["A隊比分"]) || 0;
                            const currB = parseInt(m["B隊比分"]) || 0;
                            const legA = currA - prevA;
                            const legB = currB - prevB;

                            const rowClass = mLive ? "leg-card live" : (mDone ? "leg-card done" : "leg-card");

                            legInfoHtml += `
                                <div class="${rowClass}">
                                    <div class="leg-round">${m["輪次"]}</div>
                                    <div class="leg-players">
                                        <div class="p-side">${[m["A隊員1"], m["A隊員2"], m["A隊員3"]].filter(p => p && p !== "待定").join("/")}</div>
                                        <div class="p-vs">vs</div>
                                        <div class="p-side">${[m["B隊員1"], m["B隊員2"], m["B隊員3"]].filter(p => p && p !== "待定").join("/")}</div>
                                    </div>
                                    <div class="leg-score">
                                        <span class="s-val ${legA > legB ? 'win' : ''}">${legA >= 0 ? '+'+legA : legA}</span>
                                        <span class="s-sep">:</span>
                                        <span class="s-val ${legB > legA ? 'win' : ''}">${legB >= 0 ? '+'+legB : legB}</span>
                                    </div>
                                </div>
                            `;
                            prevA = currA;
                            prevB = currB;
                        });

                        if (this.isBroadcastMode) {
                            //電視轉播模式渲染
                            chasingContainer.classList.add("broadcast-grid");
                            
                            chasingContainer.innerHTML += `
                                <div class="broadcast-card animate-fadeIn" style="border-left-color: ${areaColor};">
                                    <div class="card-top-bar">
                                        <div style="display:flex; align-items:center; gap:10px;">
                                            <span class="badge-pill gray" style="font-size:0.8rem; padding:2px 10px;">${this.getChasingSequenceLabel(matches, "序號 ")}</span>
                                            <span style="color:${areaColor}; font-weight:900; font-size:1.1rem; letter-spacing:1px;">${first["區"]}</span>
                                        </div>
                                        <div style="display:flex; align-items:center; gap:10px;">
                                            <span style="color:var(--text-dim); font-size:0.9rem;"><i class="fas fa-desktop"></i> ${first["場地"]} 場</span>
                                            ${statusHtml}
                                        </div>
                                    </div>
                                    
                                    <div class="broadcast-score-row">
                                        <div class="team-block">
                                            <div class="broadcast-team-name">${first["A隊名"]}</div>
                                        </div>
                                        
                                        <div class="broadcast-score-wrapper">
                                            <div class="broadcast-score">${aScore}</div>
                                            <div class="broadcast-divider">:</div>
                                            <div class="broadcast-score">${bScore}</div>
                                        </div>

                                        <div class="team-block">
                                            <div class="broadcast-team-name">${first["B隊名"]}</div>
                                        </div>
                                    </div>

                                    ${isRelay ? `
                                    <div class="broadcast-legs-list">
                                        ${legInfoHtml}
                                    </div>
                                    ` : ''}

                                    <div class="broadcast-footer">
                                        <i class="fas fa-user-tie"></i> 裁判: ${last["裁判"] || "尚未指派"} | <i class="far fa-clock"></i> ${first["比賽時間"] || ""}
                                    </div>
                                </div>
                            `;

                        } else {
                            // 標準卡片模式
                            chasingContainer.classList.remove("broadcast-grid");
                            chasingContainer.innerHTML += `
                                <div class="card match-card animate-fadeIn" style="padding: 1.2rem; margin-bottom: 1.5rem; ${borderStyle}">
                                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom: 0.8rem; padding-bottom: 0.5rem; border-bottom: 1px solid rgba(255,255,255,0.1);">
                                        <span style="color:var(--text-dim); font-size: 0.9rem;"><i class="fas fa-running"></i> ${this.getChasingSequenceLabel(matches)}</span>
                                        <span style="color:${areaColor}; font-weight:bold;">${first["區"]}</span>
                                        ${statusHtml}
                                    </div>
                                    <div style="color: var(--text-main); font-size: 0.85rem; margin-bottom:0.8rem; font-weight:500;">
                                        ${first["區"].includes("賽") ? "" : "追分戰 - "}${first["場地"]}場
                                    </div>
                                    <div style="display:flex; justify-content:space-between; align-items:center; gap: 10px; margin-bottom:1.2rem;">
                                        <div style="text-align:center; flex:1;">
                                            <div style="font-weight:bold; font-size:1.1rem; color:white;">${first["A隊名"] || ""}</div>
                                            <div style="font-size:0.8rem; color:var(--text-dim); text-align:center;">
                                                ${this.getTeamRoster(first, "A")}
                                            </div>
                                        </div>
                                        <div style="font-size:1.8rem; font-weight:bold; color:var(--accent); min-width: 80px; text-align:center; padding: 5px 10px; background:rgba(0,0,0,0.3); border-radius:8px;">
                                            ${aScore} : ${bScore}
                                        </div>
                                        <div style="text-align:center; flex:1;">
                                            <div style="font-weight:bold; font-size:1.1rem; color:white;">${first["B隊名"] || ""}</div>
                                            <div style="font-size:0.8rem; color:var(--text-dim); text-align:center;">
                                                ${this.getTeamRoster(first, "B")}
                                            </div>
                                        </div>
                                    </div>
                                    
                                    ${isRelay ? `
                                    <div class="standard-legs-list">
                                        ${legInfoHtml}
                                    </div>
                                    ` : ''}
                                    
                                    ${last["裁判"] ? `<div style="text-align:right; margin-top:0.8rem; font-size:0.85rem; color:var(--text-dim); opacity:0.6;">裁判: ${last["裁判"]}</div>` : ''}
                                </div>
                            `;
                        }

                    });
                }
            }
        } catch(e) {
            if(rrContainer) rrContainer.innerHTML = "<div class='card' style='color:red;'>載入失敗</div>";
            if(chasingContainer) chasingContainer.innerHTML = "<div class='card' style='color:red;'>載入失敗</div>";
        }
    },

    initHistory() {
        const btn = document.getElementById("btn-v-history-search");
        if(btn && !btn.hasAttribute("data-bound")) {
            btn.setAttribute("data-bound", "true");
            btn.addEventListener("click", () => this.searchHistory());
        }
    },

    async searchHistory() {
        const query = document.getElementById("v-history-search").value.trim();
        const resultsDiv = document.getElementById("v-history-results");
        if (!query) {
            alert("請輸入姓名！");
            return;
        }

        resultsDiv.innerHTML = "<div style='text-align:center; padding: 2rem;'><i class='fas fa-spinner fa-spin fa-2x'></i><br>搜尋中...</div>";

        try {
            const rrData = await API.getSchedule("");
            const chData = await API.getChasingSchedule("");
            let matches = [];

            if (rrData && rrData.data) {
                rrData.data.forEach(m => {
                    const matchString = JSON.stringify(m);
                    if (matchString.includes(query)) {
                        m.MatchType = "預賽";
                        matches.push(m);
                    }
                });
            }

            if (chData && chData.data) {
                chData.data.forEach(m => {
                    const matchString = JSON.stringify(m);
                    if (matchString.includes(query)) {
                        m.MatchType = m["區"] && m["區"].includes("賽") ? "決賽" : "追分賽";
                        matches.push(m);
                    }
                });
            }

            if (matches.length === 0) {
                resultsDiv.innerHTML = `<div class="card" style="text-align:center; color:var(--text-dim);">找不到符合「${query}」的比賽紀錄。</div>`;
                return;
            }

            let html = `<h3 style="margin-bottom:1rem; text-align:center; color:var(--primary);">🎯 「${query}」的比賽紀錄</h3>`;

            // 分離預賽 vs 追分/淘汰賽
            const rrMatches = matches.filter(m => m.MatchType === "預賽");
            const chMatches = matches.filter(m => m.MatchType !== "預賽");

            // === 預賽：簡易卡片 ===
            if (rrMatches.length > 0) {
                html += `<h4 style="color:var(--text-dim); margin:1rem 0 0.5rem;">📋 預賽紀錄</h4>`;
                html += `<div class="team-grid">`;
                rrMatches.forEach(m => {
                    const status = m["比賽狀態"] || "待賽";
                    const isDone = status.includes("完賽");
                    const isTeamA = (m["A隊員1"] === query || m["A隊員2"] === query || m["A隊員3"] === query || m["A隊名"] === query);
                    const queryAStyle = isTeamA ? "color:var(--accent); font-weight:bold;" : "";
                    const queryBStyle = !isTeamA ? "color:var(--accent); font-weight:bold;" : "";

                    html += `
                    <div class="card" style="padding:1rem;">
                        <div style="display:flex; justify-content:space-between; border-bottom:1px solid var(--border); padding-bottom:0.5rem; margin-bottom:0.5rem;">
                            <span style="color:var(--primary);"><i class="far fa-calendar-alt"></i> ${m["年月"] || ""} | ${m.MatchType} - ${m["輪次"]}</span>
                            <span class="status-badge ${isDone ? 'status-done' : 'status-pending'}">${status}</span>
                        </div>
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <div style="text-align:center; flex:1; ${queryAStyle}">
                                <div>${m["A隊名"]}</div>
                                <small>${[m["A隊員1"], m["A隊員2"], m["A隊員3"]].filter(p => p && p !== "待定").join(", ")}</small>
                            </div>
                            <div style="padding: 0 1rem; font-size:1.5rem; font-weight:bold;">
                                ${m["A隊比分"]||0} - ${m["B隊比分"]||0}
                            </div>
                            <div style="text-align:center; flex:1; ${queryBStyle}">
                                <div>${m["B隊名"]}</div>
                                <small>${[m["B隊員1"], m["B隊員2"], m["B隊員3"]].filter(p => p && p !== "待定").join(", ")}</small>
                            </div>
                        </div>
                    </div>`;
                });
                html += `</div>`;
            }

            // === 追分/淘汰賽：與賽程查詢相同的卡片樣式 ===
            if (chMatches.length > 0) {
                html += `<h4 style="color:var(--text-dim); margin:1.5rem 0 0.5rem;">⚡ 追分/淘汰賽紀錄</h4>`;
                
                this.groupChasingMatches(chMatches).forEach(gMatches => {
                    const first = gMatches[0];
                    const last = gMatches[gMatches.length - 1];

                    const status = last["比賽狀態"] || "待賽";
                    const isDone = status.includes("完賽");
                    const isLive = status.includes("進行中");
                    const statusHtml = isDone ? `<span class="status-badge status-done">已完賽</span>` : (isLive ? `<span class="status-badge status-live" style="background:#ff4757; color:white; animation: pulse 1.5s infinite;">即時比分</span>` : `<span class="status-badge status-pending">${status}</span>`);

                    const aScore = last["A隊比分"] || 0;
                    const bScore = last["B隊比分"] || 0;

                    const areaColor = first["區"] && first["區"].includes("猛禽") ? "var(--raptor)" : (first["區"] && first["區"].includes("小鳥") ? "var(--birdie)" : "var(--accent)");
                    const borderStyle = `border-left: 4px solid ${areaColor};`;

                    // 產生接力明細
                    let legInfo = "";
                    let prevA = 0;
                    let prevB = 0;
                    const isRelay = this.isChasingSeries(gMatches);

                    if (isRelay) {
                        gMatches.forEach(m => {
                            const mStatus = m["比賽狀態"] || "待賽";
                            const mDone = mStatus.includes("完賽");
                            const mLive = mStatus.includes("進行中");
                            const rowStyle = mLive ? "background:rgba(255,107,53,0.15); color:white; font-weight:bold;" : (mDone ? "opacity:0.8;" : "opacity:0.5;");

                            const currA = parseInt(m["A隊比分"]) || 0;
                            const currB = parseInt(m["B隊比分"]) || 0;
                            const legA = currA - prevA;
                            const legB = currB - prevB;

                            legInfo += `
                                <div style="display:grid; grid-template-columns: 80px 1fr 60px 1fr 60px; gap:5px; font-size:0.75rem; padding:6px 8px; border-radius:4px; margin-top:2px; ${rowStyle}">
                                    <span style="opacity:0.9;">${m["輪次"]}</span>
                                    <span style="overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${[m["A隊員1"], m["A隊員2"], m["A隊員3"]].filter(p => p && p !== "待定").join("/")}</span>
                                    <span style="text-align:right;"><span style="color:var(--accent);">${legA >= 0 ? '+'+legA : legA}</span> <small>(${currA})</small></span>
                                    <span style="overflow:hidden; text-overflow:ellipsis; white-space:nowrap; text-align:right;">${[m["B隊員1"], m["B隊員2"], m["B隊員3"]].filter(p => p && p !== "待定").join("/")}</span>
                                    <span style="text-align:right;"><span style="color:var(--accent);">${legB >= 0 ? '+'+legB : legB}</span> <small>(${currB})</small></span>
                                </div>
                            `;
                            prevA = currA;
                            prevB = currB;
                        });
                    }

                    const legDetailHtml = isRelay ? `
                        <div style="background:rgba(0,0,0,0.2); border-radius:8px; padding:8px; border:1px solid rgba(255,255,255,0.05);">
                            <div style="display:grid; grid-template-columns: 80px 1fr 60px 1fr 60px; gap:5px; font-size:0.65rem; color:var(--text-dim); padding:0 8px 4px 8px; border-bottom:1px solid rgba(255,255,255,0.05);">
                                <span>目標</span><span>TEAM A</span><span>取分(總分)</span><span style="text-align:right;">TEAM B</span><span style="text-align:right;">取分(總分)</span>
                            </div>
                            ${legInfo}
                        </div>
                    ` : "";

                    html += `
                        <div class="card match-card animate-fadeIn" style="padding: 1.2rem; margin-bottom: 1.5rem; ${borderStyle}">
                            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom: 0.8rem; padding-bottom: 0.5rem; border-bottom: 1px solid rgba(255,255,255,0.1);">
                                <span style="color:var(--text-dim); font-size: 0.9rem;"><i class="far fa-calendar-alt"></i> ${first["年月"] || "未知日期"}${gMatches.length ? ` | ${this.getChasingSequenceLabel(gMatches)}` : ""}</span>
                                <span style="color:${areaColor}; font-weight:bold;">${first["區"]}</span>
                                ${statusHtml}
                            </div>
                            <div style="color: var(--text-main); font-size: 0.85rem; margin-bottom:0.8rem; font-weight:500;">
                                ${first["區"].includes("賽") ? "" : "追分戰 - "}${first["場地"]}場
                            </div>
                            <div style="display:flex; justify-content:space-between; align-items:center; gap: 10px; margin-bottom:1.2rem;">
                                <div style="text-align:center; flex:1;">
                                    <div style="font-weight:bold; font-size:1.1rem; color:white;">${first["A隊名"] || ""}</div>
                                    <div style="font-size:0.8rem; color:var(--text-dim); text-align:center;">
                                        ${this.getTeamRoster(first, "A")}
                                    </div>
                                </div>
                                <div style="font-size:1.8rem; font-weight:bold; color:var(--accent); min-width: 80px; text-align:center; padding: 5px 10px; background:rgba(0,0,0,0.3); border-radius:8px;">
                                    ${aScore} : ${bScore}
                                </div>
                                <div style="text-align:center; flex:1;">
                                    <div style="font-weight:bold; font-size:1.1rem; color:white;">${first["B隊名"] || ""}</div>
                                    <div style="font-size:0.8rem; color:var(--text-dim); text-align:center;">
                                        ${this.getTeamRoster(first, "B")}
                                    </div>
                                </div>
                            </div>
                            
                            ${legDetailHtml}
                            
                            ${last["裁判"] ? `<div style="text-align:right; margin-top:0.8rem; font-size:0.85rem; color:var(--text-dim); opacity:0.6;">裁判: ${last["裁判"]}</div>` : ''}
                        </div>
                    `;
                });
            }

            resultsDiv.innerHTML = html;

        } catch (e) {
            resultsDiv.innerHTML = "<div style='text-align:center; color:red;'>查詢發生錯誤。</div>";
        }
    },

    async loadPoints() {
        const grid = document.getElementById("v-points-grid");
        if (!grid) return;
        
        // 使用 Skeleton Loading 佔位符
        grid.innerHTML = Array(6).fill(0).map(() => `
            <div class="rank-card" style="opacity: 0.3; filter: grayscale(1);">
                <div class="avatar-circle" style="background: #334155;"></div>
                <div class="rank-info">
                    <div style="height: 20px; width: 100px; background: #334155; border-radius: 4px; margin-bottom: 8px;"></div>
                    <div style="height: 14px; width: 150px; background: #334155; border-radius: 4px;"></div>
                </div>
            </div>
        `).join("");
        
        try {
            // 並發請求：抓取本期存檔積分、與球員照片 Mapping (改用 getPointsRecords，不重新計算，速度極快)
            const [ptsRes, infoRes] = await Promise.all([
                API.getPointsRecords(),
                API.getPlayersInfo()
            ]);

            if (ptsRes && ptsRes.data) {
                grid.innerHTML = "";
                const photoMap = (infoRes && infoRes.status === "success") ? infoRes.data : {};
                const defaultAvatar = "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCIgZmlsbD0iIzY0NzQ4YiI+PHBhdGggZD0iTTEyIDEyYzIuMjEgMCA0LTEuNzkgNC00cy0xLjc5LTQtNC00LTQgMS43OS00IDQgMS43OSA0IDQgNHptMCAyYy0yLjY3IDAtOCAxLjM0LTggNHYyaDE2di0yYzAtMi42Ni01LjMzLTQtOC00eiIvPjwvc3ZnPg==";

                const topPlayers = ptsRes.data;
                if (topPlayers.length === 0) {
                    grid.innerHTML = "<div class='card' style='grid-column: 1 / -1; text-align:center;'>尚未執行月結統計，目前暫無資料</div>";
                    return;
                }

                // 效能優化：分批渲染 (Batch Rendering)
                const BATCH_SIZE = 15;
                let renderedCount = 0;

                const renderBatch = () => {
                    const nextBatch = topPlayers.slice(renderedCount, renderedCount + BATCH_SIZE);
                    let htmlList = "";
                    
                    nextBatch.forEach((p, index) => {
                        const actualIdx = renderedCount + index;
                        const rankNum = actualIdx + 1; // 簡化排名邏輯，加速渲染
                        
                        const rankClass = rankNum <= 3 ? `rank-${rankNum}` : "";
                        const badgeHtml = rankNum <= 3 
                            ? `<div class="rank-number"><i class="fas fa-crown" style="margin-right:4px;"></i> NO.${rankNum}</div>` 
                            : `<div class="rank-number" style="background:rgba(255,255,255,0.1);">NO.${rankNum}</div>`;
                        
                        let avatarUrl = photoMap[p.name] || defaultAvatar;
                        if (avatarUrl.includes("drive.google.com/uc?export=view&id=")) {
                            const fileId = avatarUrl.split("id=")[1];
                            avatarUrl = `https://drive.google.com/thumbnail?id=${fileId}&sz=w500`;
                        }

                        htmlList += `
                            <div class="rank-card ${rankClass} animate-fadeIn">
                                ${badgeHtml}
                                <img src="${avatarUrl}" class="avatar-circle" loading="lazy" alt="${p.name}">
                                <div class="rank-info">
                                    <h3>${p.name}</h3>
                                    <p>${p.team || "自由球員"} | ${p.area || "未分區"}</p>
                                </div>
                                <div class="rank-score">${p.totalPts}</div>
                            </div>
                        `;
                    });

                    const tempDiv = document.createElement("div");
                    tempDiv.innerHTML = htmlList;
                    while (tempDiv.firstChild) {
                        grid.appendChild(tempDiv.firstChild);
                    }

                    renderedCount += BATCH_SIZE;
                    if (renderedCount < topPlayers.length) {
                        requestAnimationFrame(renderBatch); // 讓瀏覽器在下一幀空檔繼續畫
                    }
                };

                renderBatch();
            }
        } catch(e) {
            console.error(e);
            grid.innerHTML = "<div class='card' style='grid-column: 1 / -1; text-align:center; color:red;'>讀取失敗，請稍後再試</div>";
        }
    },

    async loadSpecial() {
        const container = document.getElementById("v-special-container");
        if (!container) return;
        container.innerHTML = "<div style='text-align:center;'><i class='fas fa-spinner fa-spin'></i> 載入中...</div>";

        try {
            const res = await API.getSpecialRecords();
            if (res && res.status === "success" && res.data) {
                // 取得目前選取的年份 (預設為今年)
                const currentYM = document.getElementById("current-date")?.value || new Date().toISOString().substring(0, 10);
                const currentYear = currentYM.substring(0, 4);
                
                // 篩選出今年度的所有紀錄
                let displayData = res.data.filter(r => {
                    const rDate = r["年月"] || "";
                    return rDate.startsWith(currentYear);
                });
                
                // 按日期降序排序 (最新的在上面)
                displayData.sort((a, b) => new Date(b["年月"] || 0) - new Date(a["年月"] || 0));

                if (displayData.length === 0) {
                    container.innerHTML = `<div class='card' style='text-align:center;'>${currentYear} 年度尚無發布公告。</div>`;
                    return;
                }

                let html = `<div style="display: flex; flex-direction: column; gap: 1.5rem; width: 100%;">`;
                displayData.forEach(r => {
                    const content = r["公佈內容"] || `${r["類型"]}: ${r["姓名"]} ${r["備註"] || ""}`;
                    const date = r["年月"] || "賽事公告";
                    
                    html += `
                    <div class="card animate-fadeIn" style="padding:1.5rem; text-align:left; border-left: 5px solid var(--primary); background: rgba(59, 130, 246, 0.05); box-shadow: 0 4px 15px rgba(0,0,0,0.2);">
                        <div style="margin-bottom: 1rem; color: var(--accent); font-weight: bold; border-bottom: 1px solid rgba(255,255,255,0.1); padding-bottom: 0.5rem; display: flex; justify-content: space-between; align-items: center;">
                            <span><i class="fas fa-bullhorn"></i> ${date} 賽事公佈欄</span>
                            <span style="font-size: 0.8rem; background: var(--primary); color: white; padding: 2px 8px; border-radius: 20px; opacity: 0.8;"># 年度紀錄</span>
                        </div>
                        <div style="white-space: pre-wrap; font-size: 1.1rem; line-height: 1.8; color: white;">${content}</div>
                    </div>
                    `;
                });
                html += `</div>`;
                container.innerHTML = html;
            } else {
                container.innerHTML = "<div class='card' style='text-align:center;'>目前暫無公告內容。</div>";
            }
        } catch(e) {
            console.error("Viewer loadSpecial error:", e);
            container.innerHTML = "<div style='text-align:center;'>載入錯誤</div>";
        }
    },
    async loadRegistration() {
        const container = document.getElementById("v-reg-list-container");
        if (!container) return;
        container.innerHTML = "<div style='text-align:center; padding: 2rem;'><i class='fas fa-spinner fa-spin fa-2x'></i><br>載入名單中...</div>";

        const datePicker = document.getElementById("current-date");
        const dateToUse = datePicker ? datePicker.value : null;

        try {
            const res = await API.getRegistrations(dateToUse);
            if (res && res.status === "success") {
                this.renderRegistration(res.data);
            } else {
                this.renderRegistration([]);
            }
        } catch (e) {
            console.error("Viewer loadRegistration error:", e);
            container.innerHTML = "<div class='card' style='text-align:center; color:red;'>讀取失敗，請稍後再試</div>";
        }
    },

    renderRegistration(data) {
        const container = document.getElementById("v-reg-list-container");
        if (!container || !data || data.length === 0) {
            container.innerHTML = `
                <div class="card" style="text-align: center; color: var(--text-dim); padding: 3rem;">
                    <i class="fas fa-users-slash" style="font-size: 3rem; margin-bottom: 1rem; display: block; opacity: 0.3;"></i>
                    <p>目前尚無報名分組資料。</p>
                </div>`;
            return;
        }

        const dataTeams = [...new Set(data.map(p => String(p.隊名 || "").trim()).filter(t => t))];
        const dataAreas = [...new Set(data.map(p => String(p.區 || p.區別 || "").replace("區", "").trim()).filter(a => a))];
        
        const areas = dataAreas.length > 0 ? dataAreas : CONFIG.AREAS.map(a => a.replace("區", ""));
        const teams = dataTeams.length > 0 ? dataTeams : CONFIG.TEAMS;

        let html = `
            <div class="card animate-fadeIn" style="overflow-x: auto; padding: 0;">
                <table class="matrix-table">
                    <thead>
                        <tr>
                            <th style="background: rgba(255,255,255,0.05); color: var(--accent);">隊名 \\ 區</th>
                            ${areas.map(area => `<th>${area}</th>`).join("")}
                        </tr>
                    </thead>
                    <tbody>
        `;

        teams.forEach(team => {
            const teamColor = CONFIG.TEAM_COLORS[team] || "var(--text-main)";
            html += `
                <tr>
                    <td style="color: ${teamColor}; font-weight: bold; background: rgba(0,0,0,0.1);">
                        ${team}
                    </td>
            `;

            areas.forEach(area => {
                const cleanArea = area.replace("區", "");
                const cellPlayers = data.filter(p => {
                    const pTeam = String(p.隊名 || "").trim();
                    const pArea = String(p.區 || p.區別 || "").replace("區", "");
                    return pTeam === team && pArea === cleanArea;
                });

                html += `
                    <td>
                        <div class="player-stack">
                            ${cellPlayers.length > 0 ? 
                                cellPlayers.map(p => `<div class="p-name">${p.姓名}</div>`).join("") : 
                                "<span style='opacity:0.2'>-</span>"
                            }
                        </div>
                    </td>
                `;
            });
            html += `</tr>`;
        });

        html += `</tbody></table></div>`;

        // 顯示「未分配」人員
        const unassigned = data.filter(p => !p.隊名 || p.隊名.trim() === "");
        if (unassigned.length > 0) {
            html += `
                <div class="card animate-fadeIn" style="margin-top: 1.5rem; border: 1px dashed var(--border);">
                    <h4 style="color: var(--text-dim); margin-bottom: 0.8rem;"><i class="fas fa-clock"></i> 尚未分組人員 (${unassigned.length})</h4>
                    <div style="display: flex; flex-wrap: wrap; gap: 8px;">
                        ${unassigned.map(p => `<span class="player-tag">${p.姓名} (${(p.區 || "").replace("區", "")})</span>`).join("")}
                    </div>
                </div>
            `;
        }

        container.innerHTML = html;
    }
};
