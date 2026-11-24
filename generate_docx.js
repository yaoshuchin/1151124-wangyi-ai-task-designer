const fs = require('fs');
const docx = require('docx');
const {
    Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
    WidthType, BorderStyle, AlignmentType, ShadingType, Header, Footer, PageNumber, ExternalHyperlink
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

function createCalloutWithLink(title, content, linkText = null, linkUrl = null, bgColor = BG_LIGHT_BLUE, borderColor = SECONDARY_COLOR) {
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

    if (linkText && linkUrl) {
        paragraphs.push(
            new Paragraph({
                spacing: { before: 80, after: 40 },
                children: [
                    new TextRun({
                        text: "🌐 線上實體系統展示：",
                        bold: true,
                        size: 20,
                        font: "Microsoft JhengHei",
                        color: SECONDARY_COLOR
                    }),
                    new ExternalHyperlink({
                        children: [
                            new TextRun({
                                text: linkText + " (點擊開啟連結)",
                                style: "Hyperlink",
                                color: "0284C7",
                                underline: {},
                                size: 20,
                                font: "Microsoft JhengHei"
                            })
                        ],
                        link: linkUrl
                    })
                ]
            })
        );
    }

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
        size = 22,
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
                size: 32,
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
                size: 26,
                font: "Microsoft JhengHei",
                color: SECONDARY_COLOR
            })
        ]
    });
}

async function buildHandbook() {
    console.log("開始編譯高規格 Word 研習手冊（含 91 頁完整專案與超連結）...");

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
                    new Paragraph({ spacing: { before: 1000 } }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 200, after: 100 },
                        children: [
                            new TextRun({
                                text: "AI 任務設計師",
                                bold: true,
                                size: 56,
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
                                text: "讓 AI 成為工作好夥伴：自主學習與跨域專案全集",
                                bold: true,
                                size: 32,
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
                                text: "臺北市數位實驗高中（T-School）汪檍老師 91 頁研習講義與 20+ 個落地專案超連結實錄",
                                size: 24,
                                font: "Microsoft JhengHei",
                                color: MUTED_TEXT,
                                italics: true
                            })
                        ]
                    }),

                    createCalloutWithLink(
                        "📌 核心教育哲學金句",
                        "「PROMPT 不是重點，思考流程才是：要做什麼 ➔ 怎麼做 ➔ 怎麼修。」\n「我們更像一起成長的夥伴：你教我，我也教你。」\n「你要有更專業的背景與清晰的需求，才能當一個稱職的甲方。」\n「先求有，再求好，持續迭代。」",
                        null,
                        null,
                        BG_WARM,
                        ACCENT_COLOR
                    ),

                    new Paragraph({ spacing: { before: 1200 } }),

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
                                    createCell("南投國中全體教師、國中小數位學習與資訊科技推動教師", false, "FFFFFF", AlignmentType.LEFT, 70)
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("簡報版本收錄", true, BG_LIGHT_BLUE, AlignmentType.CENTER, 30),
                                    createCell("雙版本整合：91 頁最新落地版 (20261124) + 73 頁精華版", false, "FFFFFF", AlignmentType.LEFT, 70)
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("產製與備查日期", true, BG_LIGHT_BLUE, AlignmentType.CENTER, 30),
                                    createCell("中華民國 115 年 11 月 24 日（2026/11/24）", false, "FFFFFF", AlignmentType.LEFT, 70)
                                ]
                            })
                        ]
                    }),

                    new Paragraph({ pageBreakBefore: true })
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
                                        text: "AI 任務設計師：讓AI成為工作好夥伴 | 汪檍老師",
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
                                    new TextRun({ text: "第 ", size: 18, font: "Microsoft JhengHei" }),
                                    new TextRun({ children: [PageNumber.CURRENT], size: 18, font: "Microsoft JhengHei" }),
                                    new TextRun({ text: " 頁，共 ", size: 18, font: "Microsoft JhengHei" }),
                                    new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, font: "Microsoft JhengHei" }),
                                    new TextRun({ text: " 頁", size: 18, font: "Microsoft JhengHei" })
                                ]
                            })
                        ]
                    })
                },
                children: [
                    h1("特別導言：Canva 簡報與專案線上系統總導覽"),
                    p("本研習手冊完整複製並收錄汪檍老師在 2026/11/24 發表之最新簡報內容，包括所有實體線上運行之系統、專案網頁與評量工具連結："),

                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                            new TableRow({
                                children: [
                                    createHeaderCell("資源項目", 25),
                                    createHeaderCell("專案名稱與領域", 35),
                                    createHeaderCell("線上連結網址（點擊開啟）", 40)
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("最新 91 頁簡報", true, BG_LIGHT_BLUE),
                                    createCell("AI任務設計師：讓AI成為工作好夥伴 (20261124 - GPT)"),
                                    createCell("https://www.canva.com/design/DAHZ1N312-w/9RUqHr3ISjqY8geL9zzPCQ/view")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("精華 73 頁簡報", true, BG_LIGHT_BLUE),
                                    createCell("AI任務設計師：GEMINI/GPT畫布裡的自主學習與跨域專案"),
                                    createCell("https://www.canva.com/design/DAHA4t2-pWY/HtqtOGLrNZ3m4ILw3eNU2g/view")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("校園管理", true, BG_LIGHT_GRAY),
                                    createCell("TSCHOOL 空間借用系統"),
                                    createCell("https://sites.google.com/tschool.tp.edu.tw/0909/home?authuser=2")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("校園管理", true, BG_LIGHT_GRAY),
                                    createCell("TSCHOOL 學習成果平台"),
                                    createCell("https://sites.google.com/tschool.tp.edu.tw/tschool20250915/tschool")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("校園管理", true, BG_LIGHT_GRAY),
                                    createCell("TSCHOOL 招生心理測驗"),
                                    createCell("https://sites.google.com/tschool.tp.edu.tw/tschooltest?usp=sharing")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("校園管理", true, BG_LIGHT_GRAY),
                                    createCell("TSCHOOL 數位校史平臺"),
                                    createCell("https://sites.google.com/tschool.tp.edu.tw/tschool-historyt/t-school-history")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("自主學習", true, BG_LIGHT_GRAY),
                                    createCell("字音字形自學互動網"),
                                    createCell("https://sites.google.com/view/tschool100/00年字音字形競賽練習-互動學習系統")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("自主學習", true, BG_LIGHT_GRAY),
                                    createCell("學習歷程指南互動自學平台"),
                                    createCell("https://sites.google.com/view/tschoollearningportfolio/1031")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("自主學習", true, BG_LIGHT_GRAY),
                                    createCell("AI 學習歷程密室 (GitHub Pages)"),
                                    createCell("https://ivy7579-spec.github.io/ai-learning-portfolio-game/")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("自主學習", true, BG_LIGHT_GRAY),
                                    createCell("傾城之戀閱讀解謎"),
                                    createCell("https://sites.google.com/view/loveinfallencity2026/loveinfallencity")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("自主學習", true, BG_LIGHT_GRAY),
                                    createCell("2025 臺灣科教展互動闖關"),
                                    createCell("https://sites.google.com/view/tschooldigitalera/數位時代")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("自主學習", true, BG_LIGHT_GRAY),
                                    createCell("2026 臺灣科教展互動闖關 (ChatGPT Site)"),
                                    createCell("https://edtech-2026-explore.tschool-taip-6197.chatgpt.site")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("課程教學", true, BG_LIGHT_GRAY),
                                    createCell("自我覺察工作坊報告書"),
                                    createCell("https://sites.google.com/view/tschool-self-awareness/tschool-自我覺察")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("課程教學", true, BG_LIGHT_GRAY),
                                    createCell("自我覺察 115 挑戰賽 (ChatGPT Site)"),
                                    createCell("https://self-awareness-115-challenge.tschool-taip-6197.chatgpt.site")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("課程教學", true, BG_LIGHT_GRAY),
                                    createCell("跨校交流學習遊戲 Demo"),
                                    createCell("https://school-exchange-115-demo.tschool-taip-6197.chatgpt.site/game/")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("策展共創", true, BG_LIGHT_GRAY),
                                    createCell("2026 數實祭總舖師上菜 (GitHub Pages)"),
                                    createCell("https://ivy7579-spec.github.io/chef-serving-game/?v=f74554c")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("策展共創", true, BG_LIGHT_GRAY),
                                    createCell("2026 數實祭官方活動網頁"),
                                    createCell("https://sites.google.com/view/tschool2026/2026數實祭")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("生活跨域", true, BG_LIGHT_GRAY),
                                    createCell("出國自由行協作記帳網頁"),
                                    createCell("https://sites.google.com/view/trip-count/出國自由行協作記帳網頁")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("生活跨域", true, BG_LIGHT_GRAY),
                                    createCell("臺灣轉型正義小測驗"),
                                    createCell("https://sites.google.com/view/transitional-justice-yoyo/臺灣轉型正義小測驗")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("生活跨域", true, BG_LIGHT_GRAY),
                                    createCell("易經六十四卦自學網"),
                                    createCell("https://sites.google.com/view/yichin20260310/64掛自學網?authuser=0")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("回饋評量", true, BG_LIGHT_GRAY),
                                    createCell("NPDL 前後測 Google 試算表"),
                                    createCell("https://docs.google.com/spreadsheets/d/1tDBWxZFrzeEbKllCir8JUL6GVM7id9ee/")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("回饋評量", true, BG_LIGHT_GRAY),
                                    createCell("研習回饋與反思 Google 試算表"),
                                    createCell("https://docs.google.com/spreadsheets/d/1o8b6GZ5cQDbrK3NzoNjdgnWWy9aZLlonI6wqp1V3KZY/")
                                ]
                            })
                        ]
                    }),

                    h1("第一章：淺談與 AI 的工作心理歷程與認知轉變"),
                    p("在生成式 AI 爆發的時代，許多教育工作者最初面對 AI 時往往經歷了「焦慮、抗拒、懷疑」的心態。然而，汪檍老師指出：AI 已經成為行政、教學、生活與心情的重要夥伴，從官方邀請信、課程學習單、冰箱剩菜料理、策展主題發想，到冗長繁複的 N 種計畫報告，無所不包。"),

                    h2("1.1 態度與心態的三重轉變"),
                    p("• 使用態度：過去將其視為單向索取答案的「工具」；現在將其視為雙向成長的「真人夥伴」。\n• 心態轉變：過去面對技術更迭感到「抗拒與焦慮」；現在建立客觀專業的「判斷與審核力」。\n• 深層感受：放掉 SOP，什麼面向都可以嘗試；AI 把屬於我的寶貴時間交還給我，堆疊出深厚的心理韌性。"),

                    h2("1.2 評估啟動 AI 專案的好時機"),
                    p("1. 高度客製化：現有工具無法滿足學校獨特的評量或跨域需求。\n2. 人手資源不足：時間、專業、經費匱乏，由單兵教師獨立作戰。\n3. 需求規範清晰：專案目標非常明確，清楚想要什麼成果畫面。"),

                    h1("第二章：Vibe Coding 理念與協作心法"),
                    p("「Vibe Coding」是以自然語言引導、快速原型驗證（MVP）的沉浸式協作模式。重點在於「探索可能性」而非長期維護，使用者往往就是開發者本人，試錯成本極低。"),

                    h2("2.1 與 AI 協作的碎念心理準備"),
                    p("• 如同任何專案一樣，一定要預留工作時程，事情永遠不會像你想的那麼順。\n• AI 其實沒那麼聰明，給出越清楚的架構，專案推進越流暢。\n• 要有大顆的心臟，拼命砍掉重來是協作日常，適時學會「轉彎找替代方案」。"),

                    h1("第三章：和 GEMINI / GPT 搏鬥的 20+ 個真實落地專案庫實錄"),
                    p("本章完整收錄 91 頁簡報中汪檍老師在一線實作的代表性專案，並附帶真實線上運行系統網址："),

                    createCalloutWithLink(
                        "專案 1：TSCHOOL 空間借用系統",
                        "• 需求情境：新設學校空間資源調度繁複，紙本與日曆極易撞期，無程式背景教師需自力建置。\n• 踩坑與解法：放棄大型關聯資料庫，以 Google Sheet + Apps Script 打造純前端預約防呆網頁，嵌入 Google Sites 零成本運作。\n• 合格成效：全校師生可即時查看空檔登記，行政調度負擔大減。",
                        "開啟 TSCHOOL 空間借用系統",
                        "https://sites.google.com/tschool.tp.edu.tw/0909/home?authuser=2",
                        BG_LIGHT_BLUE,
                        PRIMARY_COLOR
                    ),

                    createCalloutWithLink(
                        "專案 2：TSCHOOL 學生自主學習歷程展示平台",
                        "• 需求情境：學生多元成果（影片、PDF、3D 模型、連結）分散，缺乏具設計感的專屬櫥窗。\n• 踩坑與解法：模組化分類按鍵與卡片過濾器，結合 Google Drive 雲端連結，打造典雅展示牆。\n• 合格成效：成為全校學生代表作展示標準平台，獲得訪視委員高度評價。",
                        "開啟 TSCHOOL 學習成果平台",
                        "https://sites.google.com/tschool.tp.edu.tw/tschool20250915/tschool",
                        BG_LIGHT_GRAY,
                        SECONDARY_COLOR
                    ),

                    createCalloutWithLink(
                        "專案 3：TSCHOOL 招生宣傳互動心理測驗",
                        "• 需求情境：招生博覽會上紙本折頁乏人問津，需要在 2 分鐘內吸引國中畢業生與家長。\n• 踩坑與解法：以五大核心能力對應選項，設定加權分數演算法，自動產製設計感滿滿的個人化學習光譜卡。\n• 合格成效：現場大排長龍，單日吸引超過千人互動體驗。",
                        "開啟 TSCHOOL 招生測驗",
                        "https://sites.google.com/tschool.tp.edu.tw/tschooltest?usp=sharing",
                        BG_LIGHT_BLUE,
                        PRIMARY_COLOR
                    ),

                    createCalloutWithLink(
                        "專案 4：TSCHOOL 數位校史平臺",
                        "• 需求情境：創校邁向第五年，大量影音與文字活動紀錄分散，需要能每年持續更新並外掛於官網的平臺。\n• 踩坑與解法：運用分類鍵清楚分類，建置影音牆，以 Google Sheet 部屬達成前後台無縫同步。\n• 合格成效：每年影音資料即時匯入，省下重新委外建置官網的高額費用。",
                        "開啟 TSCHOOL 校史平臺",
                        "https://sites.google.com/tschool.tp.edu.tw/tschool-historyt/t-school-history",
                        BG_LIGHT_GRAY,
                        SECONDARY_COLOR
                    ),

                    createCalloutWithLink(
                        "專案 5：字音字形自主練字與手寫辨識網",
                        "• 需求情境：導生國文寫作別字率高，紙本練習追蹤困難，高度依賴平板，需要手寫與下載歷程記錄。\n• 踩坑與解法：逐次餵題庫導致「吃了後面丟了前面」；果斷拆解模組化分頁，加大畫布框支援手指流暢書寫。\n• 合格成效：選手可離線自主練習，系統精確記錄錯字，備賽效率倍增。",
                        "開啟字音字形自學網",
                        "https://sites.google.com/view/tschool100/00年字音字形競賽練習-互動學習系統",
                        BG_WARM,
                        ACCENT_COLOR
                    ),

                    createCalloutWithLink(
                        "專案 6：學習歷程指南互動自學平台",
                        "• 需求情境：突然受邀分享教育部發布的學習歷程指南，手冊厚重缺乏超連結與閱讀索引。\n• 踩坑與解法：將靜態長文重構為「四步驟互動目錄」，附帶檢核清單與常見誤區，大幅優化電子閱覽體驗。\n• 合格成效：成為學生撰寫學習歷程前不可或缺的前導自學教材。",
                        "開啟學習歷程指南平台",
                        "https://sites.google.com/view/tschoollearningportfolio/1031",
                        BG_LIGHT_BLUE,
                        PRIMARY_COLOR
                    ),

                    createCalloutWithLink(
                        "專案 7：2025 與 2026 臺灣科教展互動闖關系統",
                        "• 需求情境：高一校必修「數位時代」年度學習活動，開展當天半天內必須完工發布給全體學生。\n• 踩坑與解法：2025 年歷經文字超出框架、iOS 畫面回彈、PDF 匯出失敗等崩潰，放棄完美小步修復；2026 年進化導入 AI Agent，2 小時內即高效優化完成。\n• 合格成效：全展期數百名學生以手機流暢闖關完成學習單繳交。",
                        "開啟 2026 科教展闖關系統 (ChatGPT Site)",
                        "https://edtech-2026-explore.tschool-taip-6197.chatgpt.site",
                        BG_LIGHT_GRAY,
                        SECONDARY_COLOR
                    ),

                    createCalloutWithLink(
                        "專案 8：AI 學習歷程密室逃脫遊戲",
                        "• 需求情境：初次嘗試將生成式 AI 結合密室逃脫機制，以遊戲化方式帶領學生理解學習歷程的核心內涵。\n• 踩坑與解法：克服網頁狀態保存難題，將謎題線索與學習歷程檔案指標相互鑲嵌，支援即時驗證解鎖密碼。\n• 合格成效：引爆學生自發闖關熱潮，並開源部署於 GitHub Pages 供全台教師使用。",
                        "開啟 AI 學習歷程密室 (GitHub Pages)",
                        "https://ivy7579-spec.github.io/ai-learning-portfolio-game/",
                        BG_WARM,
                        ACCENT_COLOR
                    ),

                    createCalloutWithLink(
                        "專案 9：傾城之戀閱讀解謎遊戲",
                        "• 需求情境：國文文學經典張愛玲《傾城之戀》閱讀教學，擺脫單向講述，讓學生化身偵探深入文本意象。\n• 踩坑與解法：結合文本細節線索與情境謎題，引導學生在段落字裡行間尋找破案密碼。\n• 合格成效：學生對經典文本的重讀率與詮釋深度顯著提升。",
                        "開啟傾城之戀解謎",
                        "https://sites.google.com/view/loveinfallencity2026/loveinfallencity",
                        BG_LIGHT_BLUE,
                        PRIMARY_COLOR
                    ),

                    createCalloutWithLink(
                        "專案 10：自我覺察工作坊報告書與挑戰賽",
                        "• 需求情境：兩日跨域工作坊需結合學生實體作品照片進行反思復盤，落實社會情緒學習（SEL）指標。\n• 踩坑與解法：設計生活化引導問題與五項度雷達圖，打造結合 ChatGPT Site 的線上挑戰互動介面。\n• 合格成效：學生產出圖文並茂的深度反思報告，直接收錄於學習歷程檔案。",
                        "開啟自我覺察 115 挑戰賽",
                        "https://self-awareness-115-challenge.tschool-taip-6197.chatgpt.site",
                        BG_LIGHT_GRAY,
                        SECONDARY_COLOR
                    ),

                    createCalloutWithLink(
                        "專案 11：2026 數實祭總舖師上菜-數位辦桌趣",
                        "• 需求情境：全校大型策展需要將台灣「辦桌」意象轉化為師生共創的數位體驗。\n• 踩坑與解法：史詩級崩潰！「AI 手把手教串接 API，串好才發現每點一次都要燒錢！」果斷轉彎改用純前端 Canvas 拼貼。\n• 合格成效：全展期逾 2,000 人次參與共創數位佳餚，無伺服器與 API 負擔！",
                        "開啟總舖師上菜共創遊戲 (GitHub Pages)",
                        "https://ivy7579-spec.github.io/chef-serving-game/?v=f74554c",
                        BG_WARM,
                        ACCENT_COLOR
                    ),

                    createCalloutWithLink(
                        "專案 12：自由行協作即時記帳與分帳平台",
                        "• 需求情境：師生海外參訪多幣別計算混亂，市售 App 無法精確符合分帳需求。\n• 踩坑與解法：在 Google Sheet 先建架構，前端做大按鍵優化，最小轉帳次數演算法精確算出一覽表。\n• 合格成效：旅程中多人即時登錄花費，結算時一鍵產出清單，零爭議！",
                        "開啟自由行協作記帳",
                        "https://sites.google.com/view/trip-count/出國自由行協作記帳網頁",
                        BG_LIGHT_BLUE,
                        PRIMARY_COLOR
                    ),

                    createCalloutWithLink(
                        "專案 13：臺灣轉型正義小測驗",
                        "• 需求情境：公民與歷史跨領域教學，將嚴肅歷史議題轉化為具備啟發性與同理心的情境問答。\n• 踩坑與解法：避免偏頗斷語，以史實情境做多角度思考提問，測驗後產出引導反思卡。\n• 合格成效：課堂討論熱烈，學生能站在歷史人物立場進行價值辨析。",
                        "開啟臺灣轉型正義小測驗",
                        "https://sites.google.com/view/transitional-justice-yoyo/臺灣轉型正義小測驗",
                        BG_LIGHT_GRAY,
                        SECONDARY_COLOR
                    ),

                    createCalloutWithLink(
                        "專案 14：易經六十四卦自學網",
                        "• 需求情境：國文與哲學自主學習，讓學生藉由卦象智慧反思當前專案或生活面臨的抉擇瓶頸。\n• 踩坑與解法：打破文言玄奧，以「釐清困惑 ➔ 搖卦 ➔ 白話解讀 ➔ 行動反思」四步驟引導。\n• 合格成效：成為學生專案卡關時極佳的哲學思維復盤工具。",
                        "開啟易經六十四卦自學網",
                        "https://sites.google.com/view/yichin20260310/64掛自學網?authuser=0",
                        BG_WARM,
                        ACCENT_COLOR
                    ),

                    h1("第四章：AI 專案流程心法與教育現場的深遠意義"),
                    h2("4.1 當個稱職的甲方"),
                    p("許多人以為 AI 專案成功的關鍵是知道某個神秘的 Prompt，但汪檍老師強調：\n1. 需求明確清晰：你比 AI 更清楚最終畫面的呈現。\n2. 具備專業背景與鑑賞力：以自身專業一眼識破生成內容的破綻。\n3. 給予充足上下文：提供範例與邊界限制。\n4. 小步驗收與耐性調校：先求有，再求好，持續迭代。"),

                    h2("4.2 AI 代工 vs. 深度學習"),
                    p("「AI 讓產出變快，但快不起來的是思考的過程。」"),
                    p("如果只是讓 AI 代工寫作業、做簡報，學生的思維只會退化；但如果將 AI 作為鏡子與鷹架，引導學生經歷「發想 ➔ 提示 ➔ 審核 ➔ 判斷 ➔ 修正」，這才是真正發生深刻學習的時刻。在未來，每個孩子都能在 AI 的賦能下，客製化自己的學習畫布！"),

                    new Paragraph({ spacing: { before: 200 } }),
                    createCalloutWithLink(
                        "🌟 研習總結心得",
                        "「我們更像一起成長的夥伴，你教我，我也教你。」\n感謝汪檍老師傾囊相授 91 頁真實跌撞與豐碩成果，願每位老師都能勇敢拿起 AI 畫筆，成為自主學習與教育創新的卓越任務設計師！",
                        "汪檍老師 Facebook 專案紀錄交流",
                        "https://www.facebook.com/profile.php?id=100000512468628",
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
    console.log("✅ Word 研習手冊已重新編譯完成（含完整 91 頁專案與超連結）！");
}

buildHandbook().catch(console.error);
