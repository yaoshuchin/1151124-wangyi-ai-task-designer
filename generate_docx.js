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

function createLinkCell(linkText, targetUrl, widthPercent = null) {
    return new TableCell({
        shading: { fill: "FFFFFF", type: ShadingType.CLEAR },
        borders: thinBorder,
        children: [
            new Paragraph({
                alignment: AlignmentType.LEFT,
                spacing: { before: 80, after: 80, line: 260 },
                children: [
                    new ExternalHyperlink({
                        children: [
                            new TextRun({
                                text: linkText,
                                style: "Hyperlink",
                                color: "0284C7",
                                underline: { type: "single" },
                                size: 20,
                                font: "Microsoft JhengHei",
                                bold: true
                            })
                        ],
                        link: targetUrl
                    })
                ]
            })
        ],
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
                            top: { style: BorderStyle.NONE },
                            bottom: { style: BorderStyle.NONE },
                            right: { style: BorderStyle.NONE },
                            left: { style: BorderStyle.SINGLE, size: 24, color: borderColor }
                        },
                        children: paragraphs
                    })
                ]
            })
        ]
    });
}

function h1(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 360, after: 180 },
        children: [
            new TextRun({
                text: text,
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
        spacing: { before: 240, after: 120 },
        children: [
            new TextRun({
                text: text,
                bold: true,
                size: 26,
                font: "Microsoft JhengHei",
                color: SECONDARY_COLOR
            })
        ]
    });
}

function p(text, isBold = false) {
    return new Paragraph({
        spacing: { before: 60, after: 100, line: 300 },
        children: [
            new TextRun({
                text: text,
                size: 22,
                font: "Microsoft JhengHei",
                color: DARK_TEXT,
                bold: isBold
            })
        ]
    });
}

function bullet(text, boldPrefix = "") {
    const children = [];
    if (boldPrefix) {
        children.push(new TextRun({
            text: boldPrefix,
            bold: true,
            size: 21,
            font: "Microsoft JhengHei",
            color: PRIMARY_COLOR
        }));
    }
    children.push(new TextRun({
        text: text,
        size: 21,
        font: "Microsoft JhengHei",
        color: DARK_TEXT
    }));

    return new Paragraph({
        bullet: { level: 0 },
        spacing: { before: 40, after: 60, line: 280 },
        children: children
    });
}

// ----------------------------------------------------
// 1. Build 1151125 Teacher Handbook
// ----------------------------------------------------
async function build1151125Handbook() {
    console.log("Generating 1151125 Handbook...");
    const doc = new Document({
        styles: {
            default: {
                document: {
                    run: { font: "Microsoft JhengHei", size: 22, color: DARK_TEXT },
                    paragraph: { spacing: { line: 300 } }
                }
            }
        },
        sections: [
            {
                properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
                headers: {
                    default: new Header({
                        children: [
                            new Paragraph({
                                alignment: AlignmentType.RIGHT,
                                children: [
                                    new TextRun({
                                        text: "AI賦能：製作動態視覺化ＡＰＰ · 實務手冊 (1151125 講師版)",
                                        size: 16,
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
                                    new TextRun({ text: "第 ", size: 18, color: MUTED_TEXT, font: "Microsoft JhengHei" }),
                                    new TextRun({ children: [PageNumber.CURRENT], size: 18, color: MUTED_TEXT }),
                                    new TextRun({ text: " 頁，共 ", size: 18, color: MUTED_TEXT, font: "Microsoft JhengHei" }),
                                    new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: MUTED_TEXT }),
                                    new TextRun({ text: " 頁", size: 18, color: MUTED_TEXT, font: "Microsoft JhengHei" })
                                ]
                            })
                        ]
                    })
                },
                children: [
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 200, after: 100 },
                        children: [
                            new TextRun({
                                text: "AI賦能：製作教學使用動態視覺化多模態 APP",
                                bold: true,
                                size: 36,
                                font: "Microsoft JhengHei",
                                color: PRIMARY_COLOR
                            })
                        ]
                    }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 60, after: 200 },
                        children: [
                            new TextRun({
                                text: "臺師大教創中心陳佩英教授團隊 × 優履團隊 · 完整研習備查與實務指引手冊",
                                size: 24,
                                font: "Microsoft JhengHei",
                                color: SECONDARY_COLOR,
                                bold: true
                            })
                        ]
                    }),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                            new TableRow({
                                children: [
                                    createHeaderCell("項目", 25),
                                    createHeaderCell("研習詳細資訊", 75)
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("課程主題", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("AI賦能: 製作動態視覺化ＡＰＰ（1151125 講師版）")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("指導與研發來源", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("國立臺灣師範大學 教育創新與評鑑研究中心（陳佩英教授） 與 優履團隊")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("主講與協作陣容", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("主講值日生：王琴雲 老師\n協作團隊：江英瑞、李佳穎、陳惠珍、黎懿瑩 老師")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("研習日期", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("民國 115 年 11 月 25 日（2026-11-25）")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("原版 Canva 簡報", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createLinkCell("Canva 53 頁完整講綱簡報連結", "https://canva.link/ssgcc8l72qtigkt")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("核心目標", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("1. 熟悉 Gemini Canvas 環境與操作模式\n2. 培養「教學需求」轉化為「AI 工具」的設計思維\n3. 開發學科專屬動態教學 APP\n4. 建立跨域共創社群與結構化專業回饋")
                                ]
                            })
                        ]
                    }),
                    new Paragraph({ spacing: { before: 200, after: 100 }, children: [] }),
                    h1("第壹章：工作坊核心目標與分組協作架構"),
                    p("本工作坊旨在引領中學各學科教師，打破對程式開發的恐懼，學會扮演「稱職的甲方與教學設計師」。透過自然語言提示詞（Prompt）引導 Google NotebookLM 與 Gemini Canvas，將傳統枯燥的紙本學習單昇華為高互動、具即時回饋與視覺動態特質的教學 APP。"),
                    h2("一、工作坊四大核心學習目標"),
                    bullet(" 了解生成式 AI 在網頁互動元件的生成原理，熟練文字指令與畫面微調。", "1. 熟悉 Canvas 環境與操作模式："),
                    bullet(" 不依賴複雜代碼，專注於梳理「誰在用、怎麼用、出錯怎麼辦」的互動本質。", "2. 需求轉化設計思維："),
                    bullet(" 結合學科知識點，打造涵蓋正念安頓、任務抽籤、水的三態、幾何面積等專屬教具。", "3. 開發學科動態教學 APP："),
                    bullet(" 透過「我喜歡/我希望/如果可以」三面向回饋法，形成校際跨域共創生態。", "4. 跨域共創與專業回饋："),
                    h2("二、工作坊小組協作名冊（共四組）"),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                            new TableRow({
                                children: [
                                    createHeaderCell("組別", 15),
                                    createHeaderCell("協作桌長", 20),
                                    createHeaderCell("學員名單（學科領域）", 65)
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("第 1 組 (G1)", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("陳惠珍 老師", false, "FFFFFF", AlignmentType.CENTER),
                                    createCell("塗瑀真(國文)、李欣霈(國文)、蔡艷卿(國文)")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("第 2 組 (G2)", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("江英瑞 老師", false, "FFFFFF", AlignmentType.CENTER),
                                    createCell("高若喬(自然)、葉 敏(自然)、詹智傑(自然)")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("第 3 組 (G3)", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("黎懿瑩 老師", false, "FFFFFF", AlignmentType.CENTER),
                                    createCell("林福泉(資訊)、姚述勤(資訊)、張雅文(數學)、鄭鳳惠(數學)、嚴高一(數學)")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("第 4 組 (G4)", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("李佳穎 老師", false, "FFFFFF", AlignmentType.CENTER),
                                    createCell("吳芃軒(健體)、鄭燕琪(綜合)、曾麗娜(英語)")
                                ]
                            })
                        ]
                    }),
                    new Paragraph({ spacing: { before: 200, after: 100 }, children: [] }),
                    h1("第貳章：研習 9 大實體線上系統與實作卡一覽表"),
                    p("本工作坊所有實體操作均建立在真實上線的 AI 工具、指令卡平台與共創看板之上，點擊下方各項連結即可直達操作現場："),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                            new TableRow({
                                children: [
                                    createHeaderCell("序號", 10),
                                    createHeaderCell("系統 / 資源名稱", 40),
                                    createHeaderCell("平台類型 / 說明", 30),
                                    createHeaderCell("線上連結", 20)
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("1", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("1.1 SEL 正念 / 安頓 APP"),
                                    createCell("Gemini Canvas 動態 APP"),
                                    createLinkCell("前往體驗", "https://share.gemini.google/PTjqlSLxMszJ")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("2", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("1.2 班經任務分配轉盤 APP"),
                                    createCell("Gemini Canvas 動態 APP"),
                                    createLinkCell("前往體驗", "https://share.gemini.google/QU4OpI6Js5eh")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("3", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("【SEL & 班經指令卡】資源網"),
                                    createCell("Netlify 指令卡資料庫"),
                                    createLinkCell("前往資源庫", "https://sel-class-van.netlify.app")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("4", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("Google NotebookLM 工作室"),
                                    createCell("Notebook 互動網頁生成引擎"),
                                    createLinkCell("前往 NotebookLM", "https://notebook.google.com")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("5", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("3.2 水的三態偵探 APP"),
                                    createCell("自然學科互動教具"),
                                    createLinkCell("前往體驗", "https://share.gemini.google/fk2JefsNaEq8")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("6", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("3.2 面積剪剪貼貼 APP"),
                                    createCell("數學幾何虛擬教具"),
                                    createLinkCell("前往體驗", "https://share.gemini.google/qT051JVIH3vO")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("7", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("【闖關活動與虛擬教具指令卡】"),
                                    createCell("學科專屬指令卡平台"),
                                    createLinkCell("前往指令卡", "https://subjects-van.netlify.app")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("8", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("【學習目標與題庫卡片 / 題庫指令】"),
                                    createCell("教學提示詞與題庫產生庫"),
                                    createLinkCell("前往題庫庫", "https://subjects-aids-van.netlify.app")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("9", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("2.6 & 3.6 研習成果共創看板"),
                                    createCell("Padlet 交流與互評平台"),
                                    createLinkCell("前往 Padlet", "https://padlet.com/vanessa5881/115-11-25-app-s025qbrpwmopvznipg3q")
                                ]
                            })
                        ]
                    }),
                    new Paragraph({ spacing: { before: 200, after: 100 }, children: [] }),
                    h1("第參章：課堂痛點對齊與 Check-In 探究思維"),
                    p("在開工動手寫 Prompt 之前，必須精準診斷課堂痛點。許多教師花費一整晚排版精美的紙本學習單，發下去後往往只換來「要打分數嗎？」或是寫完發呆的被動反應。動態 APP 的目的不是取代教師，而是替學生創造自我導航的「學習鷹架」。"),
                    createCalloutWithLink("教學 APP 創作核心句型（Check-In 公式）",
                        "『我想做一個【課堂使用場景】使用的 APP，解決學生【學習卡關痛點】的問題，並且讓他們【獲得關鍵教學成效】。』\n" +
                        "範例：『我想做一個【學生各自用平板】使用的 APP，解決學生【答錯就放棄、卡住不知下一步】的問題，並且讓他們【獲得提示卡即時回饋，願意主動嘗試繼續往下做】。』"
                    ),
                    h2("一、課堂使用場景的四大模式"),
                    bullet(" 著重字體特大、對比強烈、單鍵全班觸控抽籤、倒數計時與動畫慶祝效果。", "1. 老師投影，帶全班一起用："),
                    bullet(" 著重每人自主配速、即時對錯回饋、關卡鎖定與成就徽章機制。", "2. 學生各自用平板操作："),
                    bullet(" 著重組員輪流輸入、角色防呆機制、組內意見投票與共識決策器。", "3. 學生分組共用一台："),
                    bullet(" 著重無教師在場時的防呆提示卡、語音發音與自我檢核回饋機制。", "4. 回家自己練習："),
                    new Paragraph({ spacing: { before: 200, after: 100 }, children: [] }),
                    h1("第肆章：階段二：SEL 與班級經營 APP 實作全攻略"),
                    p("社會情緒學習（SEL）與班級經營是課堂專注力的基石。透過轉盤抽籤機、情緒溫度計與呼吸安頓工具，能迅速凝聚學生學習動能。"),
                    h2("一、好指令的三大思考面向（寫指令就是做教學設計）"),
                    bullet(" 誰在用？使用環境是投影機大螢幕還是手機平板？操作者的年齡與先備能力為何？", "1. 【角色與背景】："),
                    bullet(" 怎麼用？按鈕觸發的流程為何？是否有進度條、計時器、音效與過關判定條件？", "2. 【互動邏輯】："),
                    bullet(" 出錯怎麼辦？防範輸入空白、負數人數、惡意狂點按鈕導致程式當機之教學規則。", "3. 【防呆機制】："),
                    h2("二、教學 APP 敏捷開發四步循環"),
                    bullet(" 於 NotebookLM 或 Canvas 輸入教學需求，產出初版互動網頁。", "Step 1 生成："),
                    bullet(" 拋開教師思維，假想自己是後段或調皮學生，故意亂按並操作一次。", "Step 2 試用："),
                    bullet(" 具體記錄哪裡卡住、字體太小、還是邏輯與課綱不合。", "Step 3 說出問題："),
                    bullet(" 使用自然語言與截圖指引 AI 修正，回到 Step 1 迭代，直到滿意。", "Step 4 修改指令："),
                    h2("三、三大關鍵優化原則"),
                    bullet(" 改太多就無法辨識是哪一句提示詞發揮正面作用，容易產生非預期副作用。", "原則 1：一次只改一件事 ——"),
                    bullet(" 避免「做得酷一點」，改用「按鈕放大為 48px、增加彈跳動畫與綠色陰影」。", "原則 2：描述精確具體 ——"),
                    bullet(" 遇到滿意成果先備份或存為新版本，在新分支中大膽測試進階功能。", "原则 3：善用版本分支 ——"),
                    h2("四、工具鏈三階發布工作流（零伺服器維護技巧）"),
                    p("本工作坊推薦採用「NotebookLM ➔ Gemini Canvas ➔ Google Sites」黃金整合流程："),
                    bullet(" 進入 NotebookLM，在左側新增來源貼上【SEL 指令卡】，對話框輸入：「請在右側工作室提供可預覽的互動式教學網頁」，並於右側下載產出之 HTML 檔案。", "第一階（原型生成）："),
                    bullet(" 開啟 Gemini Canvas，上傳剛下載的 HTML 檔，透過自然語言與視覺截圖進行精緻化修整。", "第二階（精緻打磨）："),
                    bullet(" 點擊 Canvas 程式碼「匯出」，直接連動 Google Sites 協作平台以嵌入式元件發布。若日後需要更新 APP，只需在協作平台後台替換程式碼，學生端之網址與 QR Code 完全無須重新印製！", "第三階（協作平台發布）："),
                    new Paragraph({ spacing: { before: 200, after: 100 }, children: [] }),
                    h1("第伍章：階段三：學科領域與虛擬教具 APP 實作"),
                    p("學科教學的核心在於將「抽象概念」視覺化。透過「三個面向 ＋ 一條輸送帶」架構，各學科教師皆能高效產出專屬教具。"),
                    h2("一、「三個面向 ＋ 一條輸送帶」架構思維"),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                            new TableRow({
                                children: [
                                    createHeaderCell("工作站 / 模組", 25),
                                    createHeaderCell("核心關注焦點", 35),
                                    createHeaderCell("提示詞指令範例", 40)
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("工作站 1：核心\n(結構與邏輯)", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("定義知識點、計分規則、題庫抽籤與過關條件。"),
                                    createCell("「設計一個三態水分子模擬器，加熱時分子運動加快，冷卻時凝固結晶。」")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("工作站 2：外觀\n(視覺與風格)", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("介面美化、配色主題、字體對比與動態微動畫。"),
                                    createCell("「採用溫暖木質調背景，卡片具備圓角 20px 與翻牌 3D 旋轉特效。」")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("工作站 3：例外\n(防呆與進階)", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("防連按、極值保護、無答案時之預設提醒。"),
                                    createCell("「學生未選擇答案時停用『下一題』按鈕，提交前跳出二次確認彈窗。」")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("輸送帶：全程適用\n(全域格式規範)", true, BG_WARM, AlignmentType.CENTER),
                                    createCell("程式製作規範、單一 HTML 輸出、跨裝置相容性。"),
                                    createCell("「請以單一純 HTML/CSS/JS 交付，響應式排版，引用 Tailwind CSS 與 Lucide 圖標。」")
                                ]
                            })
                        ]
                    }),
                    h2("二、APP 上線前三道把關原則（教師的專業判斷是 AI 無法取代的）"),
                    bullet(" 課本用詞是否精準？例如「分子」、「原子」不可混淆，數學符號需符合教育部標準。", "1. 內容正確："),
                    bullet(" 嚴禁在 APP 中要求學生填寫真實姓名、身分證字號或聯絡資訊，一律採座號或代號。", "2. 學生隱私："),
                    bullet(" 中小學階段嚴禁要求學生自行註冊外部付費或高風險 AI 帳號，全數以封裝好之前端運行。", "3. 年齡適切："),
                    h2("三、動態 APP 自評檢核表（五大指標）"),
                    bullet(" 開場設定的 2～3 項教學痛點，APP 確實完全解決。", "指標 1【做到選定的需求】："),
                    bullet(" 具備動態視覺、立即回饋或操作實驗性，確實勝過印刷紙本。", "指標 2【勝過紙本學習單】："),
                    bullet(" 題目、概念與用詞皆對照課本標準，無幻覺或偏誤。", "指標 3【學科品質把關】："),
                    bullet(" 介面直覺，不需教師長篇解說，學生摸索 10 秒即可自行操作。", "指標 4【介面親和易懂】："),
                    bullet(" 連續按、空白重來都不會當機崩潰，容錯率極高。", "指標 5【強健防呆耐操】："),
                    h2("四、多模態英語文 APP 實戰案例：【生活單字圖鑑 APP】"),
                    p("本案例結合 Gemini 多模態影像分析與文字轉語音（TTS）積木："),
                    bullet(" 學生用平板拍下身邊物件（如蘋果、自行車、路牌），APP 自動辨識物品英文，並以標準母語口音朗讀單字與例句，建立強烈視聽連結。", "情境應用："),
                    bullet(" 「你是一位充滿熱情的兒童英語老師。步驟 1：分析使用者上傳之相片主體。步驟 2：乾淨輸出該物品『英文單字』、『KK音標』與『極簡例句（限 5 字以內）』。步驟 3：輸出乾淨文字格式供 TTS 語音朗讀。」", "驅動 Prompt："),
                    new Paragraph({ spacing: { before: 200, after: 100 }, children: [] }),
                    h1("第陸章：專業回饋、課堂復盤與反思交流"),
                    p("在工作坊最後階段，各組將成果上傳至專屬 Padlet 看板，並進行跨領域結構化專業回饋。"),
                    h2("一、小組同儕結構化專業回饋三句型（三擇一）"),
                    bullet(" 「例：我喜歡這個抽籤機彈跳的動畫，視覺效果非常吸睛，能瞬間吸引學生注意力！」", "【我喜歡...】："),
                    bullet(" 「例：我希望『重新開始』的按鈕可以放大一些，在學生使用小尺寸平板時更好點擊。」", "【我希望...】："),
                    bullet(" 「例：如果這個計時器在最後 5 秒能加上『心跳聲』的音效或視覺閃爍，闖關感會更強烈！」", "【如果可以...】："),
                    h2("二、研習總結復盤兩大核心反思"),
                    p("Q1：回想早上第一張投影片的「紙本學習單」。經過今天的實作，您覺得這個「動態 APP」能幫你解決教學上的什麼痛點？"),
                    p("Q2：在剛剛的同儕交流中，其他領域老師的作品或回饋，給了您什麼跨越科目的教學新靈感？"),
                    new Paragraph({ spacing: { before: 200, after: 100 }, children: [] }),
                    createCalloutWithLink("著作與分享規範",
                        "本研習課程內容由國立臺灣師範大學教育創新與評鑑研究中心（陳佩英教授）與優履團隊研發。\n" +
                        "公開介紹、發表或推廣成果時，敬請註明出處與團隊貢獻，攜手打造良善之教育 AI 創新生態系。"
                    )
                ]
            }
        ]
    });

    const buffer = await Packer.toBuffer(doc);
    const outputPath = '1151125_AI賦能_製作教學使用多模態APP_研習手冊.docx';
    fs.writeFileSync(outputPath, buffer);
    console.log(`[1151125] Successfully generated: ${outputPath}`);
}

// ----------------------------------------------------
// 2. Build 1151124 Wang Yi Handbook
// ----------------------------------------------------
async function build1151124Handbook() {
    console.log("Generating 1151124 Handbook...");
    const linksData = JSON.parse(fs.readFileSync('all_links_clean.json', 'utf8'));

    const doc = new Document({
        styles: {
            default: {
                document: {
                    run: { font: "Microsoft JhengHei", size: 22, color: DARK_TEXT },
                    paragraph: { spacing: { line: 300 } }
                }
            }
        },
        sections: [
            {
                properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
                headers: {
                    default: new Header({
                        children: [
                            new Paragraph({
                                alignment: AlignmentType.RIGHT,
                                children: [
                                    new TextRun({
                                        text: "AI任務設計師：自主學習與跨域專案 · 研習行政與實務手冊 (汪檍老師)",
                                        size: 16,
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
                                    new TextRun({ text: "第 ", size: 18, color: MUTED_TEXT, font: "Microsoft JhengHei" }),
                                    new TextRun({ children: [PageNumber.CURRENT], size: 18, color: MUTED_TEXT }),
                                    new TextRun({ text: " 頁，共 ", size: 18, color: MUTED_TEXT, font: "Microsoft JhengHei" }),
                                    new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: MUTED_TEXT }),
                                    new TextRun({ text: " 頁", size: 18, color: MUTED_TEXT, font: "Microsoft JhengHei" })
                                ]
                            })
                        ]
                    })
                },
                children: [
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 200, after: 100 },
                        children: [
                            new TextRun({
                                text: "AI 任務設計師：GEMINI/GPT 畫布裡的自主學習與跨域專案",
                                bold: true,
                                size: 36,
                                font: "Microsoft JhengHei",
                                color: PRIMARY_COLOR
                            })
                        ]
                    }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 60, after: 200 },
                        children: [
                            new TextRun({
                                text: "臺北市數位實驗高中（T-School） 汪檍（汪汪）老師 · 完整行政與實務手冊",
                                size: 24,
                                font: "Microsoft JhengHei",
                                color: SECONDARY_COLOR,
                                bold: true
                            })
                        ]
                    }),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                            new TableRow({
                                children: [
                                    createHeaderCell("項目", 25),
                                    createHeaderCell("內容摘要", 75)
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("主講人", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("汪檍（汪汪）老師（臺北市數位實驗高中 T-School）")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("研習對象", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("南投國中全體教師、國中小數位學習與資訊應用推動教師")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("雙版本 Canva 簡報", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("1. 完整精進版 (91頁): https://www.canva.com/design/DAHZ1N312-w/...\n2. 核心精華版 (73頁): https://www.canva.com/design/DAHA4t2-pWY/...")
                                ]
                            }),
                            new TableRow({
                                children: [
                                    createCell("核心宗旨", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell("從「工具使用者」轉變成「任務發想者與稱職的甲方」；以自然語言引導（Vibe Coding）在真實教學中落地。")
                                ]
                            })
                        ]
                    }),
                    new Paragraph({ spacing: { before: 200, after: 100 }, children: [] }),
                    h1("第一章：研習背景與核心理念"),
                    p("生成式 AI 正在重塑教育現場。汪檍老師透過近百個真實專案探索，提出「AI 任務設計師」的理念：教學者不再只是學習單的產製者，而是任務的情境架構師。"),
                    h2("1.1 研習核心金句與思考框架"),
                    bullet("「PROMPT 不是重點，思考流程才是：要做什麼 ➔ 怎麼做 ➔ 怎麼修。」", "核心心法 1："),
                    bullet("「我們更像一起成長的夥伴：你教我，我也教你。」", "核心心法 2："),
                    bullet("「你要有更專業的背景與需求鑑賞力，才能當一個稱職的甲方。」", "核心心法 3："),
                    bullet("「先求有，再求好，持續碎念式迭代。」", "核心心法 4："),
                    bullet("「AI 讓產出變快，但快不起來的是：思考的過程（這正是教師與學生的核心價值所在）。」", "核心心法 5："),
                    new Paragraph({ spacing: { before: 200, after: 100 }, children: [] }),
                    h1("第二章：汪檍老師實作專案與線上展廳一覽（26 處真實落地站點）"),
                    p("本研習收錄簡報中所有真實上線運作的 Google Sites、GitHub Pages、ChatGPT Web App 與雲端試算表："),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                            new TableRow({
                                children: [
                                    createHeaderCell("序號", 10),
                                    createHeaderCell("專案 / 系統名稱", 45),
                                    createHeaderCell("類型 / 架構", 25),
                                    createHeaderCell("線上連結", 20)
                                ]
                            }),
                            ...linksData.map((item, idx) => new TableRow({
                                children: [
                                    createCell((idx + 1).toString(), true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                    createCell(getProjectTitle(item.url)),
                                    createCell(getProjectPlatform(item.url)),
                                    createLinkCell("前往專案", item.url)
                                ]
                            }))
                        ]
                    }),
                    new Paragraph({ spacing: { before: 200, after: 100 }, children: [] }),
                    h1("第三章：16 個深度落地專案實務剖析"),
                    p("汪檍老師在 91 頁簡報中詳盡公開了跨足行政、教學、展演與生活的 16 個真實專案，剖析從「崩潰踩坑」到「成功落地」的完整歷程："),
                    createCalloutWithLink(
                        "專案 1：TSCHOOL 空間借用系統",
                        "• 需求情境：創校初期空間多變，傳統登記撞期混亂，急需低門檻空間管理系統。\n• 踩坑與解法：Google Sites 介面結合 Google Sheets + Apps Script，前後端無縫同步。\n• 合格成效：成為全校師生每日必用之穩定行政基礎設施！",
                        "開啟 TSCHOOL 空間借用系統",
                        "https://sites.google.com/tschool.tp.edu.tw/0909/home?authuser=2",
                        BG_LIGHT_BLUE,
                        PRIMARY_COLOR
                    ),
                    createCalloutWithLink(
                        "專案 5：2026 臺灣科教展現場互動闖關系統",
                        "• 需求情境：開展前數小時需一套無痛容納數百人手機掃碼闖關的系統。\n• 踩坑與解法：極限 Vibe Coding，無資料庫負擔的純前端百分比進度條與通關彩帶。\n• 合格成效：開展當天順暢抗住大量人流，展現 AI 敏捷開發極致價值！",
                        "開啟 2026 科教展互動闖關",
                        "https://edtech-2026-explore.tschool-taip-6197.chatgpt.site",
                        BG_WARM,
                        ACCENT_COLOR
                    ),
                    createCalloutWithLink(
                        "專案 6：AI 學習歷程密室逃脫遊戲",
                        "• 需求情境：初次嘗試將生成式 AI 結合密室逃脫機制，以遊戲化方式帶領學生理解學習歷程的核心內涵。\n• 踩坑與解法：克服網頁狀態保存難題，將謎題線索與學習歷程檔案指標相互鑲嵌，支援即時驗證解鎖密碼。\n• 合格成效：引爆學生自發闖關熱潮，並開源部署於 GitHub Pages 供全台教師使用。",
                        "開啟 AI 學習歷程密室逃脫",
                        "https://ivy7579-spec.github.io/ai-learning-portfolio-game/",
                        BG_LIGHT_BLUE,
                        SECONDARY_COLOR
                    ),
                    createCalloutWithLink(
                        "專案 11：2026 數實祭總舖師上菜-數位辦桌趣",
                        "• 需求情境：學校年度成果盛典「數實祭」，將傳統辦桌文化轉化為師生共創的數位體驗。\n• 踩坑與解法：史詩級崩潰！「AI 手把手教串接 API，串好才發現每點一次都要燒錢！」果斷轉彎改用純前端 Canvas 拼貼。\n• 合格成效：全展期逾 2,000 人次參與共創數位佳餚，無伺服器與 API 負擔！",
                        "開啟總舖師上菜共創遊戲 (GitHub Pages)",
                        "https://ivy7579-spec.github.io/chef-serving-game/?v=f74554c",
                        BG_WARM,
                        ACCENT_COLOR
                    ),
                    new Paragraph({ spacing: { before: 200, after: 100 }, children: [] }),
                    h1("第四章：AI 專案流程心法與教育現場的深遠意義"),
                    p("許多人以為 AI 專案成功的關鍵是知道某個神秘的 Prompt，但汪檍老師強調：\n1. 需求明確清晰：你比 AI 更清楚最終畫面的呈現。\n2. 具備專業背景與鑑賞力：以自身專業一眼識破生成內容的破綻。\n3. 給予充足上下文：提供範例與邊界限制。\n4. 小步驗收與耐性調校：先求有，再求好，持續迭代。"),
                    p("「AI 讓產出變快，但快不起來的是思考的過程。」如果只是讓 AI 代工寫作業、做簡報，學生的思維只會退化；但如果將 AI 作為鏡子與鷹架，引導學生經歷「發想 ➔ 提示 ➔ 審核 ➔ 判斷 ➔ 修正」，這才是真正發生深刻學習的時刻！")
                ]
            }
        ]
    });

    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync("1151124汪憶-AI設計者_完整研習手冊.docx", buffer);
    fs.writeFileSync("AI任務設計師_實務架構手冊_汪檍.docx", buffer);
    console.log("[1151124] Successfully generated: 1151124汪憶-AI設計者_完整研習手冊.docx & AI任務設計師_實務架構手冊_汪檍.docx");
}

function getProjectTitle(url) {
    if (url.includes('0909')) return "TSCHOOL 空間借用系統";
    if (url.includes('tschool20250915')) return "TSCHOOL 學生自主學習成果平台";
    if (url.includes('tschool-historyt')) return "TSCHOOL 數位校史平臺";
    if (url.includes('tschooltest')) return "TSCHOOL 招生宣傳互動心理測驗";
    if (url.includes('edtech-2026')) return "2026 科教展互動闖關系統";
    if (url.includes('ai-learning-portfolio-game')) return "AI 學習歷程密室逃脫遊戲";
    if (url.includes('tschoollearningportfolio')) return "學習歷程指南自學平台";
    if (url.includes('learning-portfoliotschool2026')) return "AI 學習歷程密室 (專題頁)";
    if (url.includes('tschool100')) return "字音字形自學互動網";
    if (url.includes('tschool-self-awareness')) return "自我覺察自學導引平臺";
    if (url.includes('self-awareness-115-challenge') && url.includes('#')) return "自我覺察情境任務體驗";
    if (url.includes('self-awareness-115-challenge')) return "115 自我覺察挑戰關卡";
    if (url.includes('tschooldigitalera')) return "數位時代教學專題平臺";
    if (url.includes('tschool2026')) return "2026 數實祭策展官方平臺";
    if (url.includes('chef-serving-game')) return "2026 數實祭：總舖師上菜遊戲";
    if (url.includes('ziyou-huarui')) return "自由花蕊自主學習展示網站";
    if (url.includes('loveinfallencity')) return "傾城之戀文學閱讀解謎闖關";
    if (url.includes('school-exchange-115')) return "學校交流闖關互動模擬系統";
    if (url.includes('trip-count')) return "出國自由行協作記帳平台";
    if (url.includes('transitional-justice-yoyo')) return "臺灣轉型正義小測驗";
    if (url.includes('yichin20260310')) return "易經六十四卦自學網";
    if (url.includes('tteesstt2026')) return "校園互動測驗與遊戲化評量測試站";
    if (url.includes('二月的記持')) return "二月的記持社群策展實務";
    if (url.includes('1tDBWxZFrzeEbKllCir8JUL6GVM7id9ee')) return "共筆共創回饋協作試算表";
    if (url.includes('100000512468628')) return "汪檍老師個人專業交流專頁";
    if (url.includes('1o8b6GZ5cQDbrK3NzoNjdgnWWy9aZLlonI6wqp1V3KZY')) return "南投國中研習線上共筆簽到表";
    return "實務專案系統";
}

function getProjectPlatform(url) {
    if (url.includes('chatgpt.site')) return "ChatGPT Web App";
    if (url.includes('github.io')) return "GitHub Pages";
    if (url.includes('sites.google.com')) return "Google Sites";
    if (url.includes('docs.google.com/spreadsheets')) return "Google Sheets";
    if (url.includes('facebook.com')) return "社群紀錄 / 交流";
    return "Web 線上平台";
}

async function main() {
    console.log("=== Compiling all handbooks ===");
    await build1151125Handbook();
    await build1151124Handbook();
    console.log("=== All handbooks compiled successfully ===");
}

main().catch(console.error);
