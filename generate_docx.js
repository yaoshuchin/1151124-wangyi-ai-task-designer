const fs = require('fs');
const docx = require('docx');
const {
    Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
    WidthType, BorderStyle, AlignmentType, ShadingType, Header, Footer, PageNumber
} = docx;

// Color Palette
const PRIMARY_COLOR = "1E3A8A";   // Deep Indigo Blue
const SECONDARY_COLOR = "0284C7"; // Sky Blue
const ACCENT_COLOR = "D97706";    // Amber / Accent
const DARK_TEXT = "1E293B";       // Slate 800
const MUTED_TEXT = "475569";      // Slate 600
const BG_LIGHT_BLUE = "F0F9FF";   // Sky 50
const BG_LIGHT_GRAY = "F8FAFC";   // Slate 50
const BG_WARM = "FEF3C7";         // Amber 100
const BORDER_COLOR = "CBD5E1";    // Slate 300

const thinBorder = {
    top: { style: BorderStyle.SINGLE, size: 4, color: BORDER_COLOR },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: BORDER_COLOR },
    left: { style: BorderStyle.SINGLE, size: 4, color: BORDER_COLOR },
    right: { style: BorderStyle.SINGLE, size: 4, color: BORDER_COLOR },
};

function createHeaderCell(text, widthPercent = null) {
    return new TableCell({
        shading: { fill: PRIMARY_COLOR, type: ShadingType.CLEAR },
        borders: thinBorder,
        children: [
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 100, after: 100 },
                children: [
                    new TextRun({
                        text,
                        bold: true,
                        color: "FFFFFF",
                        size: 22,
                        font: "Microsoft JhengHei"
                    })
                ]
            })
        ],
        width: widthPercent ? { size: widthPercent, type: WidthType.PERCENTAGE } : undefined
    });
}

function createCell(text, isBold = false, bgColor = "FFFFFF", align = AlignmentType.LEFT, widthPercent = null) {
    const lines = text.split('\n');
    const paragraphs = lines.map(line => new Paragraph({
        alignment: align,
        spacing: { before: 80, after: 80, line: 260 },
        children: [
            new TextRun({
                text: line,
                bold: isBold,
                size: 20,
                font: "Microsoft JhengHei",
                color: DARK_TEXT
            })
        ]
    }));

    return new TableCell({
        shading: { fill: bgColor, type: ShadingType.CLEAR },
        borders: thinBorder,
        children: paragraphs,
        width: widthPercent ? { size: widthPercent, type: WidthType.PERCENTAGE } : undefined
    });
}

function createCallout(title, content, bgColor = BG_LIGHT_BLUE, borderColor = SECONDARY_COLOR) {
    const lines = content.split('\n');
    const paragraphs = [
        new Paragraph({
            spacing: { before: 60, after: 80 },
            children: [
                new TextRun({
                    text: title,
                    bold: true,
                    size: 22,
                    font: "Microsoft JhengHei",
                    color: PRIMARY_COLOR
                })
            ]
        }),
        ...lines.map(line => new Paragraph({
            spacing: { before: 40, after: 40, line: 260 },
            children: [
                new TextRun({
                    text: line,
                    size: 20,
                    font: "Microsoft JhengHei",
                    color: DARK_TEXT
                })
            ]
        }))
    ];

    return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        shading: { fill: bgColor, type: ShadingType.CLEAR },
                        borders: {
                            left: { style: BorderStyle.SINGLE, size: 24, color: borderColor },
                            top: { style: BorderStyle.NONE },
                            right: { style: BorderStyle.NONE },
                            bottom: { style: BorderStyle.NONE }
                        },
                        children: paragraphs
                    })
                ]
            })
        ]
    });
}

function p(text, options = {}) {
    const {
        bold = false,
        size = 22, // 11pt
        color = DARK_TEXT,
        align = AlignmentType.LEFT,
        before = 80,
        after = 80,
        line = 300
    } = options;

    return new Paragraph({
        alignment: align,
        spacing: { before, after, line },
        children: [
            new TextRun({
                text,
                bold,
                size,
                font: "Microsoft JhengHei",
                color
            })
        ]
    });
}

function h1(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.LEFT,
        spacing: { before: 360, after: 180 },
        children: [
            new TextRun({
                text,
                bold: true,
                size: 32, // 16pt
                font: "Microsoft JhengHei",
                color: PRIMARY_COLOR
            })
        ]
    });
}

function h2(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_2,
        alignment: AlignmentType.LEFT,
        spacing: { before: 260, after: 120 },
        children: [
            new TextRun({
                text,
                bold: true,
                size: 26, // 13pt
                font: "Microsoft JhengHei",
                color: SECONDARY_COLOR
            })
        ]
    });
}

function h3(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_3,
        alignment: AlignmentType.LEFT,
        spacing: { before: 180, after: 80 },
        children: [
            new TextRun({
                text,
                bold: true,
                size: 22, // 11pt
                font: "Microsoft JhengHei",
                color: ACCENT_COLOR
            })
        ]
    });
}

async function buildHandbook() {
    console.log("開始編譯高規格 Word 研習手冊...");

    const doc = new Document({
        styles: {
            default: {
                document: {
                    run: { font: "Microsoft JhengHei", size: 22, color: DARK_TEXT }
                }
            }
        },
        sections: [
            // ==================== 封面 ====================
            {
                properties: {
                    page: {
                        margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 }
                    }
                },
                children: [
                    new Paragraph({ spacing: { before: 1200 } }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 200, after: 100 },
                        children: [
                            new TextRun({
                                text: "AI 任務設計師",
                                bold: true,
                                size: 56, // 28pt
                                font: "Microsoft JhengHei",
                                color: PRIMARY_COLOR
                            })
                        ]
                    }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 80, after: 300 },
                        children: [
                            new TextRun({
                                text: "GEMINI / GPT 畫布裡的自主學習與跨域專案實踐手冊",
                                bold: true,
                                size: 32, // 16pt
                                font: "Microsoft JhengHei",
                                color: SECONDARY_COLOR
                            })
                        ]
                    }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 100, after: 600 },
                        children: [
                            new TextRun({
                                text: "從「工具使用者」轉變為「任務發想者」與「稱職的甲方」之專案實踐心法",
                                size: 24,
                                font: "Microsoft JhengHei",
                                color: MUTED_TEXT,
                                italics: true
                            })
                        ]
                    }),

                    createCallout(
                        "📌 核心教育哲學金句",
                        "「PROMPT 不是重點，思考流程才是：要做什麼 ➔ 怎麼做 ➔ 怎麼修。」\n「我們更像一起成長的夥伴：你教我，我也教你。」\n「你要有更專業的背景與清晰的需求，才能當一個稱職的甲方。」\n「先求有，再求好，持續迭代。」",
                        BG_WARM,
                        ACCENT_COLOR
                    ),

                    new Paragraph({ spacing: { before: 1400 } }),

                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                            new TableRow({
                                children: [
                                    createCell("主講專家", true, BG_LIGHT_BLUE, AlignmentType.CENTER, 30),
                                    createCell("汪檍（汪汪）老師\n臺北市數位實驗高中（T-School）教師 / 翻轉教育專欄作家", false, "FFFFFF", AlignmentType.LEFT, 70)
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("研習對象", true, BG_LIGHT_BLUE, AlignmentType.CENTER, 30),
                                    createCell("南投國中全體教師、國中小數位學習與資訊應用推動教師", false, "FFFFFF", AlignmentType.LEFT, 70)
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("研習主題", true, BG_LIGHT_BLUE, AlignmentType.CENTER, 30),
                                    createCell("AI 協作開發、Vibe Coding 理念、16 個落地教育專案、Prompt 任務工程", false, "FFFFFF", AlignmentType.LEFT, 70)
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("彙整產製日期", true, BG_LIGHT_BLUE, AlignmentType.CENTER, 30),
                                    createCell("中華民國 115 年 11 月 24 日（2026/11/24）", false, "FFFFFF", AlignmentType.LEFT, 70)
                                ]
                            })
                        ]
                    }),

                    new Paragraph({
                        pageBreakBefore: true
                    })
                ]
            },

            // ==================== 正文章節 ====================
            {
                properties: {
                    page: {
                        margin: { top: 1200, bottom: 1200, left: 1200, right: 1200 }
                    }
                },
                headers: {
                    default: new Header({
                        children: [
                            new Paragraph({
                                alignment: AlignmentType.RIGHT,
                                children: [
                                    new TextRun({
                                        text: "AI 任務設計師：自主學習與跨域專案實踐手冊 | 汪檍老師",
                                        size: 18,
                                        color: MUTED_TEXT,
                                        font: "Microsoft JhengHei"
                                    })
                                ]
                            })
                        ]
                    })
                },
                footers: {
                    default: new Footer({
                        children: [
                            new Paragraph({
                                alignment: AlignmentType.CENTER,
                                children: [
                                    new TextRun({
                                        text: "第 ",
                                        size: 18,
                                        font: "Microsoft JhengHei"
                                    }),
                                    new TextRun({
                                        children: [PageNumber.CURRENT],
                                        size: 18,
                                        font: "Microsoft JhengHei"
                                    }),
                                    new TextRun({
                                        text: " 頁，共 ",
                                        size: 18,
                                        font: "Microsoft JhengHei"
                                    }),
                                    new TextRun({
                                        children: [PageNumber.TOTAL_PAGES],
                                        size: 18,
                                        font: "Microsoft JhengHei"
                                    }),
                                    new TextRun({
                                        text: " 頁",
                                        size: 18,
                                        font: "Microsoft JhengHei"
                                    })
                                ]
                            })
                        ]
                    })
                },
                children: [
                    h1("第一章：淺談與 AI 的工作心理歷程與認知轉變"),
                    p("在生成式 AI 爆發的時代，許多教育工作者最初面對 AI 時往往經歷了「焦慮、抗拒、懷疑」的心態。然而，隨著技術演進，我們逐漸發現：AI 不再只是一個冰冷的自動化工具，而是一個能與我們對話、共同成長的「數位真人協作夥伴」"),

                    h2("1.1 態度與心態的三重轉變"),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                            new TableRow({
                                children: [
                                    createHeaderCell("階段進程", 20),
                                    createHeaderCell("過去的認知（傳統心態）", 40),
                                    createHeaderCell("現在與未來的跨越（任務設計師心態）", 40)
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("角色定位", true, BG_LIGHT_BLUE, AlignmentType.CENTER),
                                    createCell("將 AI 視為執行單一指令的「工具」（Tool），單向索取標準答案。"),
                                    createCell("將 AI 視為「協同創作者與真人夥伴」，雙向溝通、相互拋接思考。")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("心理狀態", true, BG_LIGHT_BLUE, AlignmentType.CENTER),
                                    createCell("擔心被取代的「焦慮與抗拒」，執著於尋找無所不能的神奇 Prompt。"),
                                    createCell("「包容與坦然」，接受 AI 會犯錯，並在除錯修正的過程中持續迭代。")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("工作模式", true, BG_LIGHT_BLUE, AlignmentType.CENTER),
                                    createCell("尋找標準作業程序（SOP），期待照本宣科得到一致成果。"),
                                    createCell("「放下 SOP，打開專案實驗思維」，在不確定性中探索未知與可能性。")
                                ]
                            })
                        ]
                    }),

                    h2("1.2 評估啟動 AI 專案的三大情境指標"),
                    p("汪檍老師在實踐中總結出，最適合邀請 AI 深度協作（甚至進入 Vibe Coding）的專案通常具備以下三個特質："),
                    createCallout(
                        "🔍 專案啟動檢核指標",
                        "1. 高度客製化：現成軟體或 SaaS 平台無法滿足學校獨特的評量、班級活動或跨領域需求。\n2. 人手資源不足：沒有專屬工程團隊或充裕經費，由教師或學生自行單兵作戰。\n3. 需求規範清晰：雖然不會寫程式，但對於「想要什麼功能、給誰用、流程長怎樣」有極為明確的畫面。",
                        BG_LIGHT_BLUE,
                        SECONDARY_COLOR
                    ),

                    h1("第二章：Vibe Coding 理念與協作心法"),
                    p("「Vibe Coding」是一種嶄新的軟體與專案協作模式：開發者不需深究底層語法細節，而是透過自然語言與 AI 共同營造專案的「氛圍與節奏」（Vibe），藉由快速提示、即時反饋、小步驗證來建構產品。"),

                    h2("2.1 Vibe Coding 的核心原則"),
                    p("👉 重點是「探索可能性」，而非打造百年維護的系統架構。\n👉 使用者往往等於開發者本人（或本校師生），試錯成本極低、容錯率高。\n👉 溝通流程：自然語言描繪情境 ➔ 生成原型 ➔ 點擊測試 ➔ 挑出痛點 ➔ 碎念式反饋 ➔ 即時修復。"),

                    h2("2.2 協作中的「多麼痛的領悟」與心理準備"),
                    p("專案過程中一定會遭遇挫折與卡關，汪檍老師特別提醒三項關鍵心理素質："),
                    createCallout(
                        "💡 專案搏鬥心法",
                        "• 怎麼做已經不是問題，『到底要做什麼』才是核心：AI 讓代工產出變快，但若目標模糊，產出的只是大量的數位垃圾。\n• 經費與技術有時是必要條件，適時學會「轉彎」：當某個 API 要收費或某個套件卡死時，不要死嗑，果斷尋找替代工具或簡化介面。\n• 保持好奇心與小步快跑：不要妄想一次建出完美龐然大物，先做出一個可點擊的小按鈕，再逐步增生功能。",
                        BG_WARM,
                        ACCENT_COLOR
                    ),

                    h1("第三章：GEMINI vs. GPT 畫布特質與 12 大 AI 工具光譜"),
                    p("不同的生成式模型如同性格迥異的合作夥伴，了解它們各自的長處與盲點，才能人盡其才："),

                    h2("3.1 兩大核心模型特質比對"),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                            new TableRow({
                                children: [
                                    createHeaderCell("比較維度", 20),
                                    createHeaderCell("GEMINI 的畫布特質", 40),
                                    createHeaderCell("GPT 的畫布特質", 40)
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("核心優勢", true, BG_LIGHT_BLUE, AlignmentType.CENTER),
                                    createCell("1. 跨域聯想力極強\n2. 視覺美感與多模態敏銳\n3. 具備隨時隨地的發想彈性\n4. 與 Google 生態系（Sheet, Drive, Docs）無縫銜接"),
                                    createCell("1. 語言修辭與邏輯架構極端嚴密\n2. 程式語法與除錯精準度高\n3. 系統穩定度與耐受性強\n4. 結構化資料抽取能力出色")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("限制與盲點", true, BG_LIGHT_BLUE, AlignmentType.CENTER),
                                    createCell("記憶力較短暫，多輪深入程式碼修改時容易忘記前置約束；偶有隨興偏題情況。"),
                                    createCell("版面視覺設計偏向保守標準化，需要更明確的美感指令引導。")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("最佳應用情境", true, BG_LIGHT_BLUE, AlignmentType.CENTER),
                                    createCell("創意發想、跨領域教材整合、圖像美感引導、情境解謎、視覺策展。"),
                                    createCell("程式架構規劃、資料庫語法設計、複雜評量量規制定、精確文本邏輯校驗。")
                                ]
                            })
                        ]
                    }),

                    h2("3.2 現代教師必備的 12 大類別 AI 工具光譜全景圖"),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                            new TableRow({
                                children: [
                                    createHeaderCell("類別", 15),
                                    createHeaderCell("代表性工具", 25),
                                    createHeaderCell("核心專長與教育應用情境", 60)
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("文字生成", true, BG_LIGHT_GRAY),
                                    createCell("ChatGPT, Claude, Gemini"),
                                    createCell("教案撰寫、文章改寫、概念提問、情境對話角色扮演、評語生成。")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("即時搜尋", true, BG_LIGHT_GRAY),
                                    createCell("Perplexity, Gemini Search"),
                                    createCell("即時時事查證、學術文獻溯源、最新法規與教學資源精確彙整。")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("知識庫整理", true, BG_LIGHT_GRAY),
                                    createCell("NotebookLM, Claude Projects"),
                                    createCell("上傳校本教材、課綱文件、學生研究論文，進行閉環知識問答與播客生成。")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("簡報生成", true, BG_LIGHT_GRAY),
                                    createCell("Gamma, Canva AI, Tome"),
                                    createCell("輸入教案大綱快速產出高質感視覺簡報，節省 80% 排版時間。")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("圖像生成", true, BG_LIGHT_GRAY),
                                    createCell("Midjourney, DALL-E 3, Leonardo"),
                                    createCell("歷史情境重現、課文插圖設計、活動視覺主視覺、概念視覺化。")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("影音生成", true, BG_LIGHT_GRAY),
                                    createCell("Runway, Sora, CapCut AI"),
                                    createCell("文字轉生動短影片、校園宣導短片動態字幕與自動剪輯合成。")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("語音技術", true, BG_LIGHT_GRAY),
                                    createCell("ElevenLabs, Whisper"),
                                    createCell("多國語言逼真配音、課堂演講自動轉精確逐字稿、聽力測驗音檔生成。")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("音樂音效", true, BG_LIGHT_GRAY),
                                    createCell("Suno, Udio"),
                                    createCell("輸入學科口訣即時創作成流行歌曲、微電影配樂與廣播劇音效。")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("程式開發", true, BG_LIGHT_GRAY),
                                    createCell("Cursor, Copilot, Antigravity"),
                                    createCell("全自動生成教學網頁、課堂即時抽籤小工具、自動化試卷產製。")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("數據分析", true, BG_LIGHT_GRAY),
                                    createCell("ChatGPT Data, Excel AI"),
                                    createCell("學生學業成績趨勢分析、問卷交叉比對、自動繪製專業統計圖表。")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("流程自動化", true, BG_LIGHT_GRAY),
                                    createCell("Make, Zapier, AI Agents"),
                                    createCell("表單提交自動寄信通知、研習簽到自動同步名冊、日常行政排程。")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("3D 與空間", true, BG_LIGHT_GRAY),
                                    createCell("Tripo AI, Luma, Spline"),
                                    createCell("文字生成 3D 模型、虛擬校園導覽、AR/VR 沉浸式情境教材建構。")
                                ]
                            })
                        ]
                    }),

                    h1("第四章：和 GEMINI / GPT 搏鬥的 16 個真實落地專案庫實錄"),
                    p("本章完整收錄汪檍老師在教學與行政現場實際打造的 16 個落地專案，從專案痛點、AI 解決策略、踩坑挫折到最終合格成效，提供最鮮活的一線實戰經驗："),

                    h2("【領域一：校園管理與行政專案】"),
                    createCallout(
                        "專案 1：TSCHOOL 空間借用系統",
                        "• 需求情境：新設學校空間資源彈性多變，傳統紙本或 Google 日曆容易撞期且無法設定審核規則。\n• 碰壁歷程：一開始想建置大型 SQL 資料庫，但技術過於複雜且介面超載；AI 寫了一堆無法落地的複雜後端架構。\n• 轉彎與解法：改用 Google Apps Script 串接 Google Sheet，搭配純 HTML/Tailwind 前端與 Google Sites 嵌入，零成本達成即時預約防呆。\n• 合格成效：全校師生可直接在單一網頁查看可用時段並一鍵送出登記，行政負擔下降 70%。",
                        BG_LIGHT_BLUE,
                        PRIMARY_COLOR
                    ),

                    createCallout(
                        "專案 2：TSCHOOL 學生自主學習歷程與成果展示平台",
                        "• 需求情境：學生自主學習成果分散在各處，缺乏統一且能展現多元多媒體形式的數位畫布。\n• 碰壁歷程：初期要求 AI 生成複雜的社群互動與權限系統，導致網頁載入緩慢且維護困難。\n• 轉彎與解法：化繁為簡，以模組化卡片呈現學生作品（連結、PDF、影片、3D 模型），並支援標籤快速過濾。\n• 合格成效：成為高一新生至高三畢業生展示個人代表作的標準櫥窗，深獲家長與訪視委員好評。",
                        BG_LIGHT_GRAY,
                        SECONDARY_COLOR
                    ),

                    createCallout(
                        "專案 3：TSCHOOL 招生宣傳互動心理測驗",
                        "• 需求情境：招生博覽會上傳統紙本簡介吸引力不足，需要讓國中生在 2 分鐘內了解自己的學習特質與學校契合度。\n• 碰壁歷程：測驗計分邏輯反覆出現分支衝突，AI 產生的隨機題庫在行動端瀏覽器排版跑版。\n• 轉彎與解法：固定 6 道情境式是非/選擇題，設定清晰的計分維度，完成後自動生成具備設計感與分享圖的結果卡。\n• 合格成效：博覽會現場引發排隊體驗潮，單日觸及超過 1,200 名潛在學生與家長。",
                        BG_LIGHT_BLUE,
                        PRIMARY_COLOR
                    ),

                    h2("【領域二：自主學習任務專案】"),
                    createCallout(
                        "專案 4：故宮文化實境解謎任務生成器",
                        "• 需求情境：帶領學生參觀故宮，傳統學習單流於抄寫展品說明牌，需要結合情境解謎激發探究興趣。\n• 碰壁歷程：AI 憑空捏造展廳文物位置與歷史事實（模型幻覺嚴重），提示詞不夠具體導致謎題毫無難度。\n• 轉彎與解法：預先餵入故宮官方正確展品圖文清單，限定解謎題型為「細節觀察 + 歷史線索推理 + 密碼鎖驗證」。\n• 合格成效：學生拿著平板在展廳主動搜尋器物細節，自主學習投入度顯著提升。",
                        BG_WARM,
                        ACCENT_COLOR
                    ),

                    createCallout(
                        "專案 5：2025 臺灣科教展現場互動闖關系統",
                        "• 需求情境：時間極度緊迫！展覽當天需完成即時闖關互動網頁，當晚就要發布給全體參訪學生！\n• 碰壁歷程：時間壓力破表！跟 AI 溝通時 Prompt 太過冗長，產出的介面文字一直溢出螢幕框架，iOS 裝置一度無法滾動！\n• 轉彎與解法：果斷放棄複雜特效，要求 AI 採用最穩固的純 Flexbox 響應式架構，將長文拆成多步驟彈窗提示。\n• 合格成效：當晚如期上線封測，全展期數百名學生順利流暢完成闖關任務，化解開展危機！",
                        BG_LIGHT_BLUE,
                        PRIMARY_COLOR
                    ),

                    createCallout(
                        "專案 6：學習歷程指南互動目錄與前導教材",
                        "• 需求情境：官方發布的 PDF 學習歷程指南長達數十頁，學生普遍反映難以閱讀且缺乏章節跳轉索引。\n• 碰壁歷程：直接轉成網頁時排版凌亂，文字堆疊難以抓住重點。\n• 轉彎與解法：利用 AI 將指南重構為「互動式目錄地圖」，搭配「概念自我檢核清單」與關鍵字搜尋功能。\n• 合格成效：學生可在動態指引下一步步完成自主學習計畫撰寫，成為全校公認最好用的工具型教材。",
                        BG_LIGHT_GRAY,
                        SECONDARY_COLOR
                    ),

                    createCallout(
                        "專案 7：字音字形自主練字與手寫辨識平台",
                        "• 需求情境：語文競賽選手日常需要高強度自主練習字音字形，渴望有即時手寫反饋與錯題本。\n• 碰壁歷程：超級痛點！「逐次餵題庫，但吃了後面的丟了前面的！」專案間隔幾天 AI 就把之前的題庫洗掉；手寫畫筆在 iPad 上嚴重卡頓延遲！\n• 轉彎與解法：放棄單頁一次性塞入數千題庫，改為依單元拆解成輕量化分頁 JSON；手寫畫框放大為大尺寸 Canvas，支援指尖平滑繪製。\n• 合格成效：選手可離線自主練習，系統自動記錄常錯字音字形，備賽效率倍增！",
                        BG_WARM,
                        ACCENT_COLOR
                    ),

                    h2("【領域三：課程教學與評量專案】"),
                    createCallout(
                        "專案 8：課堂互評回饋抽籤與即時接力工具",
                        "• 需求情境：公開觀課或專案發表時，需要全班多組同步互評抽籤，避免順序爭議與冷場。\n• 碰壁歷程：傳統工具缺乏防重複抽籤機制，多人同時點擊時後端資料庫頻繁死鎖。\n• 轉彎與解法：設計成純前端本機狀態儲存的輕量化轉盤抽籤，支援一鍵匯入分組名單與倒數計時器。\n• 合格成效：課堂節奏明快緊湊，學生全神貫注參與評量與同儕互饋。",
                        BG_LIGHT_BLUE,
                        PRIMARY_COLOR
                    ),

                    createCallout(
                        "專案 9：兩日實體工作坊復盤紀錄與 SEL 五項度自評系統",
                        "• 需求情境：高強度跨域工作坊結束後，需結合實體作品照片進行反思復盤，並落實社會情緒學習（SEL）指標。\n• 碰壁歷程：學生只會寫「很好玩、收穫很多」等空泛心得，無法深入內省自我情緒與團隊合作歷程。\n• 轉彎與解法：建立 SEL 5 大維度（自我覺察、自我管理、社會覺察、人際技巧、負責任決定）引導題庫，搭配拍照上傳與雷達圖生成。\n• 合格成效：產出圖文並茂的深度反思紀錄，直接作為自主學習成果檔案存檔。",
                        BG_LIGHT_GRAY,
                        SECONDARY_COLOR
                    ),

                    createCallout(
                        "專案 10：NPDL 深度學習前後測檢核表系統",
                        "• 需求情境：高一必修課每學期固定進行 NPDL 深度學習能力前後測，過去四年全用紙本自評，教師統計資料極度痛苦！\n• 碰壁歷程：曾嘗試用複雜大型問卷系統，但介面繁瑣導致學生填答意願低落且難以進行前後測個別對比。\n• 轉彎與解法：利用 AI 打造「UX 極致優化版表單」，背後僅需串接 Google Sheet，學生 3 分鐘內以滑桿完成 6 大全球素養自評。\n• 合格成效：前後測數據一秒匯出自動生成比較長條圖，教師免去輸入數百張紙本分數之苦！",
                        BG_WARM,
                        ACCENT_COLOR
                    ),

                    h2("【領域四：策展與展演專案】"),
                    createCallout(
                        "專案 11：2026 數實祭總舖師上菜-數位辦桌趣",
                        "• 需求情境：全校年度大型策展活動，希望以傳統「台灣辦桌」為意象，打造數位共創的互動展覽體驗。\n• 碰壁歷程：史詩級崩潰！「AI 手把手教我串接高級影像 API，好不容易調試成功，一測試才發現每點一次都要燒錢！」經費瞬間拉警報！\n• 轉彎與解法：體悟「經費有時是必要條件」，果斷轉彎！改用純前端預渲染素材庫與開源輕量化 Canvas 濾鏡，打造免 API 費用的互動辦桌拼盤。\n• 合格成效：展覽期間超過 2,000 人次參與共創數位菜餚，成為校園最吸睛的策展打卡亮點！",
                        BG_LIGHT_BLUE,
                        PRIMARY_COLOR
                    ),

                    h2("【領域五：生活、學習與身心靈跨域專案】"),
                    createCallout(
                        "專案 12：自由行協作即時記帳與分帳平台",
                        "• 需求情境：師生海外研習或自由行時，多幣別計算與多人分帳極易混亂，市售 App 往往無法完全符合客製化需求。\n• 碰壁歷程：行動端離線狀態下資料同步衝突，介面按鍵在小螢幕手機上容易誤觸。\n• 轉彎與解法：先在 Google Sheet 建立完備計算公式，前端介面優化為大按鈕計算機風格，並支援匯率快照。\n• 合格成效：旅程中多人即時登錄花費，結算時一鍵產出誰該給誰多少錢的清晰清單，零爭議！",
                        BG_LIGHT_GRAY,
                        SECONDARY_COLOR
                    ),

                    createCallout(
                        "專案 13：人類圖（Human Design）AI 深度解讀器",
                        "• 需求情境：用於學生自我探索與師生對話，需要依據九大能量中心、內在權威與人生角色產出深度個人化分析報告。\n• 碰壁歷程：最初全靠手動一項項輸入繁複數據，耗時且容易漏填；AI 產生的報告書論述品質時高時低。\n• 轉彎與解法：迭代升級為「圖片辨識技術」，只需上傳人類圖截圖，AI 自動解析中心通道，並套用嚴格的結構化報告框架。\n• 合格成效：產出論述品質高度穩定且具備諮商同理心的長篇解讀報告，成為輔導對話的絕佳破冰媒介。",
                        BG_WARM,
                        ACCENT_COLOR
                    ),

                    createCallout(
                        "專案 14：紫微斗數命盤深度解讀器",
                        "• 需求情境：將傳統東方命理智慧轉化為個人性格優勢、潛在盲點與生涯規劃的探討工具。\n• 碰壁歷程：命盤十二宮位星曜眾多，AI 一開始常會張冠李戴，將主星吉凶完全解讀混亂。\n• 轉彎與解法：同樣導入命盤圖片光學辨識，並在 System Prompt 鎖定宮位對應關係與正向心理學引導原則。\n• 合格成效：報告書兼具傳統星曜深度與現代生涯探索啟發，成功破除迷信，轉化為自我賦能之教材。",
                        BG_LIGHT_BLUE,
                        PRIMARY_COLOR
                    ),

                    createCallout(
                        "專案 15：易經六十四卦自主學習與情境決策平臺",
                        "• 需求情境：國文與哲學跨領域自主學習，希望學生藉由古老卦象智慧，反思當前學習瓶頸或專案抉擇。\n• 碰壁歷程：學生對文言文卦辭爻辭望而生畏，無法與真實生活產生連結。\n• 轉彎與解法：設計「先確認基本問題情境 ➔ 動態搖卦 ➔ 白話解讀 ➔ 專案策略引導」的四階流程。\n• 合格成效：成為學生自主學習過程中深思熟慮與復盤心態的特殊反思平台。",
                        BG_LIGHT_GRAY,
                        SECONDARY_COLOR
                    ),

                    createCallout(
                        "專案 16：未來 N 種待開啟的 AI 專案探索庫",
                        "• 核心意涵：從校園行政、課堂評量、展演互動到生活自我探索，AI 不是冷冰冰的代工機器，而是點燃無限創意的畫布。\n• 汪檍老師期許：每一位老師與學生只要擁有好奇心，隨時都能在自己的專業領域中開啟第 17、第 18 個全新專案！",
                        BG_WARM,
                        ACCENT_COLOR
                    ),

                    h1("第五章：AI 專案流程心法與教育現場的深遠意義"),
                    h2("5.1 表面上是下指令，背後是「當個稱職的甲方」"),
                    p("許多人以為 AI 專案成功的關鍵是知道某個神秘的 Prompt 關鍵字，但汪檍老師一針見血地指出："),
                    createCallout(
                        "👑 稱職甲方的四項修煉",
                        "1. 需求清晰明確：你必須比 AI 更清楚最終畫面長怎樣，不能只說「幫我做一個好用的系統」。\n2. 具備專業專業背景與鑑賞力：只有你擁有學科與教學專業，才能一眼看出 AI 寫的教案或程式是否及格。\n3. 提供充足的上下文（Context）：給予足夠的範例、限制條件與輸入格式，而不是讓 AI 瞎猜。\n4. 小步驗收與耐性調校：不奢求一次到位，採用「先求有 ➔ 再求好 ➔ 持續迭代」的策略。",
                        BG_LIGHT_BLUE,
                        PRIMARY_COLOR
                    ),

                    h2("5.2 思考流程：要做什麼 ➔ 怎麼做 ➔ 怎麼修"),
                    p("Prompt 真的不是重點，思考流程才是靈魂所在："),
                    p("• 【要做什麼】（What）：定義使用者、核心痛點、最小可行性產品（MVP）範圍。\n• 【怎麼做】（How）：選擇合適的 AI 夥伴（GEMINI 發想 / GPT 架構 / Cursor 開發），拆解功能步驟。\n• 【怎麼修】（Fix）：測試驗收、收集真實反饋、挑出錯誤、碎念式反饋並精準修復。"),

                    h2("5.3 AI 代工 vs. 深度學習：教育的核心價值"),
                    p("「AI 讓產出變快，但快不起來的是：『思考的過程』。」"),
                    p("如果只是讓 AI 代工寫作業、做簡報，學生的思維能力只會退化；但如果將 AI 作為鏡子與鷹架，引導學生經歷「發想 ➔ 提示 ➔ 審核 ➔ 判斷 ➔ 修正」，這才是真正發生深刻學習的時刻。在未來，每個孩子都能在 AI 的賦能下，客製化自己的學習畫布，打破年級與學科的藩籬！"),

                    new Paragraph({ spacing: { before: 200 } }),
                    createCallout(
                        "🌟 結語：準備好開始實驗了嗎？",
                        "「我們更像一起成長的夥伴，你教我，我也教你。」\n教育不是注滿一桶水，而是點燃一把火。願每位教育工作者都能勇敢拿起 AI 畫筆，成為自主學習與教育創新的卓越任務設計師！",
                        BG_WARM,
                        ACCENT_COLOR
                    )
                ]
            }
        ]
    });

    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync("1151124汪憶-AI設計者_完整研習手冊.docx", buffer);
    fs.writeFileSync("AI任務設計師_實務架構手冊_汪檍.docx", buffer);
    console.log("✅ Word 研習手冊產製完成：");
    console.log("   - 1151124汪憶-AI設計者_完整研習手冊.docx");
    console.log("   - AI任務設計師_實務架構手冊_汪檍.docx");
}

buildHandbook().catch(console.error);
