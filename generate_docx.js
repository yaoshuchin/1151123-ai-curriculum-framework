const fs = require('fs');
const docx = require('docx');
const {
    Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
    WidthType, BorderStyle, AlignmentType, ShadingType, Header, Footer, PageNumber
} = docx;

// Helper styling constants
const PRIMARY_COLOR = "1A365D";   // Deep navy blue
const SECONDARY_COLOR = "2B6CB0"; // Classic blue
const ACCENT_COLOR = "C53030";    // Warm red
const BG_LIGHT_BLUE = "EBF8FF";
const BG_LIGHT_GRAY = "F7FAFC";
const BG_WARM = "FEFCBF";
const BORDER_COLOR = "CBD5E0";

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
                children: [
                    new TextRun({
                        text,
                        bold: true,
                        color: "FFFFFF",
                        size: 22, // 11pt
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
                size: 20, // 10pt
                font: "Microsoft JhengHei",
                color: "2D3748"
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
    const contentLines = Array.isArray(content) ? content : [content];
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
                            bottom: { style: BorderStyle.NONE },
                        },
                        children: [
                            new Paragraph({
                                spacing: { before: 120, after: 80 },
                                children: [
                                    new TextRun({
                                        text: `【${title}】`,
                                        bold: true,
                                        size: 22,
                                        color: borderColor,
                                        font: "Microsoft JhengHei"
                                    })
                                ]
                            }),
                            ...contentLines.map(c => new Paragraph({
                                spacing: { before: 60, after: 80, line: 280 },
                                children: [
                                    new TextRun({
                                        text: c,
                                        size: 20,
                                        font: "Microsoft JhengHei",
                                        color: "2D3748"
                                    })
                                ]
                            }))
                        ]
                    })
                ]
            })
        ]
    });
}

async function buildDocument() {
    const doc = new Document({
        styles: {
            default: {
                document: {
                    run: {
                        font: "Microsoft JhengHei",
                        size: 22, // 11pt
                        color: "2D3748"
                    },
                    paragraph: {
                        spacing: { line: 280, before: 100, after: 100 }
                    }
                }
            }
        },
        sections: [{
            properties: {
                page: {
                    margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 }
                }
            },
            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.RIGHT,
                            children: [
                                new TextRun({
                                    text: "AI導入課程與評量——AI素養融入教學與學生學習的實務架構 | 施君潔主任",
                                    size: 16,
                                    color: "718096"
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
                                new TextRun({ text: "第 ", size: 18, color: "718096" }),
                                new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "718096" }),
                                new TextRun({ text: " 頁", size: 18, color: "718096" })
                            ]
                        })
                    ]
                })
            },
            children: [
                // Title Area
                new Paragraph({
                    heading: HeadingLevel.TITLE,
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 240, after: 120 },
                    children: [
                        new TextRun({
                            text: "AI 導入課程與評量",
                            bold: true,
                            size: 48,
                            color: PRIMARY_COLOR,
                            font: "Microsoft JhengHei"
                        })
                    ]
                }),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 200 },
                    children: [
                        new TextRun({
                            text: "AI 素養融入教學與學生學習的實務架構",
                            bold: true,
                            size: 28,
                            color: SECONDARY_COLOR,
                            font: "Microsoft JhengHei"
                        })
                    ]
                }),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 400 },
                    children: [
                        new TextRun({
                            text: "講師：南投縣史港國民小學 施君潔 主任",
                            bold: true,
                            size: 24,
                            color: "4A5568",
                            font: "Microsoft JhengHei"
                        })
                    ]
                }),

                // Callout of Core Metaphor
                createCallout(
                    "核心核心教學主張：拐杖 V.S. 鷹架",
                    [
                        "「AI是鷹架，不是拐杖。」",
                        "學習的真正證據不在於學生在AI輔助下完成多漂亮的產出（當下任務表現），而在於逐步撤除AI支援後（Reduced Support），學生是否具備真正的學習遷移力、深度思考力與自我調節能力。",
                        "先有『主動、積極的態度』 ➔ 再有『思考的方法』 ➔ 最後強化『使用工具的技巧』。"
                    ],
                    "F7FAFC",
                    ACCENT_COLOR
                ),

                new Paragraph({ spacing: { after: 300 }, children: [] }),

                // Section 1: 講師簡介與經歷
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    spacing: { before: 400, after: 180 },
                    children: [
                        new TextRun({
                            text: "壹、講師背景與專業資歷",
                            bold: true,
                            size: 32,
                            color: PRIMARY_COLOR
                        })
                    ]
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "施君潔主任具備行政領導力與十餘年教學現場經驗，長期致力於開發數位化國語文課程與推動生成式 AI 教學應用的專業知能。",
                            size: 22
                        })
                    ]
                }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: [
                        new TableRow({
                            children: [
                                createHeaderCell("類別", 25),
                                createHeaderCell("經歷與專業榮譽項目", 75),
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("現職與聯繫", true, BG_LIGHT_GRAY),
                                createCell("• 南投縣史港國小 教導主任\n• 聯絡電話：0927-133669 / 049-2931521 #30\n• 電子信箱：t09139.ntct@go.edu.tw", false)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("教育部輔導專家", true, BG_LIGHT_GRAY),
                                createCell("• 教育部國中小「因材網國語文輔導團」輔導委員 (112年-114年連續受聘)\n• 教育部數位學習輔導計畫 (中部) 輔導專家 (111年-114年連續受聘)\n• 教育部 5G 智慧學習 / BYOD & THSD 實施計畫輔導專家 (112年-114年連續受聘)\n• 教育部「創造性思維暨 21 世紀核心素養」種子講師", false)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("地方輔導職務", true, BG_LIGHT_GRAY),
                                createCell("• 南投縣數位學習輔導團 輔導員 (113-114學年度)\n• 南投縣國民教育輔導團「國語文領域」輔導員 (111-114學年度)\n• 南投縣教師專業發展實踐方案「進階專業回饋人才」", false)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("數位系列講師證書", true, BG_LIGHT_GRAY),
                                createCell("• 生成式 AI 融入學科領域教學 (B5-2) - 國語文\n• 生成式 AI 與教育應用 (B5-1)\n• A1 數位學習工作坊 (一)\n• A2 數位學習工作坊 (二) - 因材網\n• 數位教學專業師資培訓證書 (中教大與 Apple 主辦)", false)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("全國與地方獲獎", true, BG_LIGHT_GRAY),
                                createCell("• 114 年度教育部推動數位學習績優中小學人員「卓越獎」 (2025)\n• 114 年度教育部因材網 A2 績優講師「典範獎」 (2025)\n• 111 學年度教育部推動數位學習績優人員「特優」\n• 111、112 學年度教育部推動數位學習優良教學案例「特優 / 優等」\n• 110 學年度教育部推動數位學習優良教學案例「優選」\n• 112、113 學年度南投縣數位學習推動績優人員與優良教案「特優」", false)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("論文發表與專書", true, BG_LIGHT_GRAY),
                                createCell("• 《科技輔助自主學習理論與應用：「自學共學互學導學」教師行動指南》(2023, 共同作者)\n• 〈AI 平台融入國語文寫作教學〉(2023, 與中教大楊智為教授共同發表)\n• 〈「語」你同行新視界〉(2023, 教學研討會發表)\n• 〈PISA 創造性思維課程設計與融入合作問題解決之科技輔助自主學習課堂實作〉(2021)", false)
                            ]
                        })
                    ]
                }),

                new Paragraph({ spacing: { after: 300 }, children: [] }),

                // Section 2: 核心反思與實證研究
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    spacing: { before: 400, after: 180 },
                    children: [
                        new TextRun({
                            text: "貳、AI 輔助學習的核心反思與實證文獻基礎",
                            bold: true,
                            size: 32,
                            color: PRIMARY_COLOR
                        })
                    ]
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "在推動 AI 進入課堂之前，我們必須先釐清：AI 到底是在「幫助學生學習」，還是讓學生產生「依賴與後設認知怠惰」？國際權威實證研究已指出關鍵問題：",
                            size: 22
                        })
                    ]
                }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: [
                        new TableRow({
                            children: [
                                createHeaderCell("研究與學者", 28),
                                createHeaderCell("研究樣本與方法", 25),
                                createHeaderCell("關鍵實證發現與教育省思", 47),
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("Bastani 等人\n(2025)", true, BG_LIGHT_GRAY),
                                createCell("近千名高中數學學生實地實證研究"),
                                createCell("• 學生使用一般生成式 AI 練習時作答表現提升；\n• 但當移除 AI 後，學生成績反而比「未使用 AI 者」低 17%！\n• 結論：必須設有『教學護欄』的 AI 家教才能降低負面依賴效應。")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("Darvishi 等人\n(2024)", true, BG_LIGHT_GRAY),
                                createCell("10 門課程、1,625 名學生隨機對照研究"),
                                createCell("• AI 提示能顯著改善學生當下的同儕回饋表現；\n• 但移除提示後，部分學生無法維持相同表現；\n• 反映學生僅是『依賴 AI 表面提示』，並未將學習策略真正內化。")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("Fan 等人\n(2025)", true, BG_LIGHT_GRAY),
                                createCell("117 名大學生寫作實驗研究"),
                                createCell("• 使用 ChatGPT 組在文章表面分數進步較為明顯；\n• 但在『知識增長』與『學習遷移』方面未呈現顯著優勢；\n• 過度依賴 AI 會減少學生規劃、監控與評估等活動，形成【後設認知懶惰】。")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("SYNC 人機學習模型\n(Fan 等人, 2026)", true, BG_LIGHT_GRAY),
                                createCell("人機學習理論架構 (預印本)", false),
                                createCell("• 在 AI 協助下完成的高品質作品，僅代表『AI 支持下的任務表現 (Task performance)』，不必然代表持久學習；\n• 必須透過逐步減少 AI 鷹架、無 AI 重述及新情境遷移任務，檢驗是否形成獨立能力。")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("Dragan Gašević\n(香港大學講座教授, 2026)", true, BG_LIGHT_GRAY),
                                createCell("2026 數位學習國際論壇"),
                                createCell("• 主題：Building Human Capability for Productive Human-AI Interaction\n• 核心倡議：Ask what learners can do when AI support is reduced！")
                            ]
                        })
                    ]
                }),

                new Paragraph({ spacing: { after: 200 }, children: [] }),

                // Section 3: 課堂範式轉移 (Before vs After)
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    spacing: { before: 300, after: 140 },
                    children: [
                        new TextRun({
                            text: "教學典範轉移：從「工具示範」走向「課程轉譯 × 教學領導」",
                            bold: true,
                            size: 26,
                            color: SECONDARY_COLOR
                        })
                    ]
                }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: [
                        new TableRow({
                            children: [
                                createHeaderCell("傳統取向：工具示範 (Before - 拐杖)", 50),
                                createHeaderCell("深度取向：課程轉譯 × 教學領導 (After - 鷹架)", 50)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("• 教哪個平台（聚焦在工具功能）\n• 示範提示詞 (prompt) 抄用\n• 分享單次 40 分鐘的零碎活動\n• 以產出內容的表面效度判斷成效（看簡報美不美、學習單齊不齊、網頁或遊戲花俏度）", false, BG_LIGHT_GRAY),
                                createCell("• 教哪個核心能力與領域學習目標\n• 設計 AI 適切角色、鷹架與查證任務\n• 以共備觀議課、學生實質學習歷程帶動教師社群成長\n• 追蹤無 AI 遷移能力、評估公平性與治理風險", true, BG_LIGHT_BLUE)
                            ]
                        })
                    ]
                }),

                new Paragraph({ spacing: { after: 300 }, children: [] }),

                // Section 4: 教育部未來人才 AI 能力圖像與素養架構
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    spacing: { before: 400, after: 180 },
                    children: [
                        new TextRun({
                            text: "參、教育部未來人才 AI 能力圖像與中小學素養框架",
                            bold: true,
                            size: 32,
                            color: PRIMARY_COLOR
                        })
                    ]
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "教育部於 115.05.22 發布之未來人才 AI 能力圖像，明確揭示中小學學生面對 AI 時代所需具備的四大核心基石：",
                            size: 22
                        })
                    ]
                }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: [
                        new TableRow({
                            children: [
                                createHeaderCell("核心能力向度", 25),
                                createHeaderCell("國小階段 (學習起步：使用到思考)", 25),
                                createHeaderCell("國中階段 (學習深化：使用到認識)", 25),
                                createHeaderCell("高中階段 (學習精進：使用到協作)", 25)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("1. 運用 AI 學習", true, BG_LIGHT_GRAY),
                                createCell("認識及使用教育用 AI，會提問及進行多輪對話。"),
                                createCell("理解 AI 基本原理，修正 AI 產出內容，會交叉查證。"),
                                createCell("理解 AI 功能與限制，並能運用進階 AI 工具自主學習。")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("2. 負責任使用 AI", true, BG_LIGHT_GRAY),
                                createCell("知道 AI 不一定正確，會基本查證、不亂用。"),
                                createCell("正確標註 AI 來源，能辨識 AI 生成內容，不濫用。"),
                                createCell("具備數位安全與倫理素養，並正確標註 AI 協作情形。")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("3. 運用 AI 解決問題", true, BG_LIGHT_GRAY),
                                createCell("透過與 AI 互動，輔助完成指定學習任務。"),
                                createCell("能判斷適合用 AI 解決的問題，可以整合 AI 產出的內容。"),
                                createCell("能架構問題本質，並能與 AI 進行人機協作以解決問題。")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("4. AI 無法取代的能力", true, BG_LIGHT_GRAY),
                                createCell("發現問題、能表達創意想法、與同儕合作。"),
                                createCell("展現使用 AI 的主動性與主導性。"),
                                createCell("能展現同理心與創造力，進行獨立自主判斷。")
                            ]
                        })
                    ]
                }),

                new Paragraph({ spacing: { after: 200 }, children: [] }),

                // Section 5: 教師與學生 AI 素養能力總表
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    spacing: { before: 300, after: 140 },
                    children: [
                        new TextRun({
                            text: "臺灣中小學教師 AI 素養能力總表 (以人為中心：確保專業能動性與人機協作)",
                            bold: true,
                            size: 26,
                            color: SECONDARY_COLOR
                        })
                    ]
                }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: [
                        new TableRow({
                            children: [
                                createHeaderCell("素養維度", 18),
                                createHeaderCell("Step 1 取得理解 (Acquire/Understand)", 27),
                                createHeaderCell("Step 2 深化應用 (Deepen/Apply)", 27),
                                createHeaderCell("Step 3 創造轉型 (Create/Transform)", 28)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("倫\nAI 倫理", true, BG_LIGHT_GRAY),
                                createCell("【掌握AI技術與倫理】\n✓ 理解AI運作與倫理原則\n✓ 保護個資與AI數據隱私\n✓ 辨識AI風險與安全防範"),
                                createCell("【內化AI倫理教學責任】\n✓ 探究AI倫理責任歸屬\n✓ 深化AI安全實踐應用\n✓ 健全校園AI使用規範"),
                                createCell("【形塑以AI促進創新的倫理】\n✓ 融入倫理思辨於AI問題解決教學\n✓ 創新AI倫理教學模式\n✓ 共構師生AI倫理實踐準則")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("技\nAI 基礎與技術", true, BG_LIGHT_GRAY),
                                createCell("【培養AI的基礎知識】\n✓ 區分AI的定義與類別\n✓ 理解AI的數據與運作\n✓ 識別AI的侷限與風險"),
                                createCell("【應用AI的技能】\n✓ 選用適性適境之AI工具\n✓ 優化提示以實踐AI協作\n✓ 執行監督以檢核AI產出"),
                                createCell("【使用AI創新】\n✓ 建構AI自動化管理教育工作\n✓ 提升數據優化以精準AI生成\n✓ 參與AI校園治理部署與決策")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("教\nAI 賦能教學", true, BG_LIGHT_GRAY),
                                createCell("【善用AI輔助學科教學】\n✓ 評估AI工具輔助教學的適用性\n✓ 分析AI工具輔助教學的有效性\n✓ 反思與優化AI工具輔助教學應用"),
                                createCell("【整合AI與教學法】\n✓ 評估AI與教學法整合適切性\n✓ 設計AI與教學法整合策略\n✓ 反思與改進教學法與AI工具效益"),
                                createCell("【促進AI教學轉型創新】\n✓ 客製AI教學工具與資源\n✓ 創新AI教學與設計學習體驗\n✓ 促進教師與AI協作的正向循環")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("發\nAI 促進教師專業發展", true, BG_LIGHT_GRAY),
                                createCell("【覺知AI時代角色與終身學習】\n✓ 理解AI時代教師自主責任\n✓ 評估教師AI素養規劃增能\n✓ 辨識與選用促進發展的數據"),
                                createCell("【AI與數據分析驅動發展】\n✓ 參與社群共學反思教師角色\n✓ 運用AI分析教學數據規劃發展\n✓ 善用AI模擬工具支持專業發展"),
                                createCell("【結合AI共創專業發展】\n✓ 重塑教學創新建立AI協作指引\n✓ 活用人機協作轉化專業實踐\n✓ 創建AI驅動的組織學習文化")
                            ]
                        })
                    ]
                }),

                new Paragraph({ spacing: { after: 200 }, children: [] }),

                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    spacing: { before: 300, after: 140 },
                    children: [
                        new TextRun({
                            text: "臺灣中小學生 AI 素養能力總表 (以人為中心：捍衛學習者能動性與形塑未來)",
                            bold: true,
                            size: 26,
                            color: SECONDARY_COLOR
                        })
                    ]
                }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: [
                        new TableRow({
                            children: [
                                createHeaderCell("素養維度", 18),
                                createHeaderCell("Step 1 取得理解 (Acquire/Understand)", 27),
                                createHeaderCell("Step 2 深化應用 (Deepen/Apply)", 27),
                                createHeaderCell("Step 3 創造轉型 (Create/Transform)", 28)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("倫\nAI 倫理", true, BG_LIGHT_GRAY),
                                createCell("【理解AI倫理困境】\n✓ 具備對AI的批判思考\n✓ 思辨AI中的應用與取捨\n✓ 覺察AI倫理展現人類價值"),
                                createCell("【熟練正用AI】\n✓ 落實AI使用規範與責任\n✓ 正用AI符合學術誠信\n✓ 堅持AI的隱私保護"),
                                createCell("【養成AI技術與倫理】\n✓ 構思AI設計倫理內涵\n✓ 建構與AI協作設計倫理\n✓ 深化AI設計倫理為習慣")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("技\nAI 基礎與技術", true, BG_LIGHT_GRAY),
                                createCell("【認知AI基礎以參與互動】\n✓ 理解AI運作原理與限制\n✓ 連結AI服務與資料來源\n✓ 覺察AI的世界並產生好奇心"),
                                createCell("【管理AI工具多元情境應用】\n✓ 實踐AI強化運算思維\n✓ 運用提示詞優化AI互動\n✓ 調整與AI協作的適應力"),
                                createCell("【創造以人為中心的AI專案】\n✓ 構想AI人文方案\n✓ 展現同理心驅動AI專案創新\n✓ 活用AI解決在地問題")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("學\nAI 賦能學習", true, BG_LIGHT_GRAY),
                                createCell("【理解AI在學習的角色界線】\n✓ 理解AI輔助學習的角色功能\n✓ 辨識適合個人學習的AI工具\n✓ 認知AI可優化學習提升動機"),
                                createCell("【應用AI提升領域學習】\n✓ 掌握AI促進領域學習策略\n✓ 應用AI工具強化領域學習\n✓ 展現AI強化學習的適應力"),
                                createCell("【開創人機協作賦能自主學習】\n✓ 構想AI協作自主學習計畫\n✓ 轉化自我、共同與社會調節\n✓ 建立AI賦能的成長思維")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("設\nAI 系統設計", true, BG_LIGHT_GRAY),
                                createCell("【理解AI系統設計原理】\n✓ 理解AI系統設計的核心\n✓ 熟悉AI系統設計的方法\n✓ 洞悉AI系統的設計責任"),
                                createCell("【應用設計思維於AI開發】\n✓ 整合理論與實踐深化AI系統\n✓ 善用工具協作提升開發效能\n✓ 管理AI系統設計專案確保品質"),
                                createCell("【創造與優化AI系統】\n✓ 設計跨領域AI系統解決複雜問題\n✓ 迭代中優化AI系統功能\n✓ 確保AI系統提升人類福祉")
                            ]
                        })
                    ]
                }),

                new Paragraph({ spacing: { after: 300 }, children: [] }),

                // Section 6: 六個提問與課堂六步實踐
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    spacing: { before: 400, after: 180 },
                    children: [
                        new TextRun({
                            text: "肆、六個提問判斷 AI 要不要進課堂",
                            bold: true,
                            size: 32,
                            color: PRIMARY_COLOR
                        })
                    ]
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "在備課與設計學習活動時，教師可透過以下六大關鍵提問，自我檢視是否真的需要導入 AI：",
                            size: 22
                        })
                    ]
                }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: [
                        new TableRow({
                            children: [
                                createHeaderCell("編號", 12),
                                createHeaderCell("檢核提問", 30),
                                createHeaderCell("深度檢核意涵與設計指標", 58)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("01", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                createCell("學什麼？", true),
                                createCell("學生必須【獨立理解】或【能做出的領域目標】是什麼？目標是否清晰具體？")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("02", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                createCell("為何用 AI？", true),
                                createCell("AI 增加何種回饋、比較、模擬或表達機會？是否真能突破傳統教學限制？")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("03", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                createCell("誰負責思考？", true),
                                createCell("哪些關鍵思考步驟必須由【學生先想、選擇、解釋與修正】，不可被 AI 掠奪？")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("04", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                createCell("如何查證？", true),
                                createCell("要引導學生用哪些課本、實驗、權威資料庫或多重來源來交叉比對 AI 的回答？")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("05", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                createCell("如何揭露？", true),
                                createCell("學生要保留哪些提問歷程、AI 回應、採納取捨與後續改寫修訂的紀錄？")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("06", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                createCell("如何證明學會？", true),
                                createCell("【在不用 AI 時】，學生能否透過口說表達、紙筆作答或在新情境中獨立完成任務？")
                            ]
                        })
                    ]
                }),

                new Paragraph({ spacing: { after: 200 }, children: [] }),

                // Section 7: 學生課堂六部曲流程
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    spacing: { before: 300, after: 140 },
                    children: [
                        new TextRun({
                            text: "學生課堂流程：先保留思考，再讓 AI 成為可被質疑的夥伴",
                            bold: true,
                            size: 26,
                            color: SECONDARY_COLOR
                        })
                    ]
                }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: [
                        new TableRow({
                            children: [
                                createHeaderCell("步驟", 15),
                                createHeaderCell("核心動作", 20),
                                createHeaderCell("學生具體任務", 40),
                                createHeaderCell("對應教師引導提示語", 25)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("第一步", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                createCell("1. 先想", true),
                                createCell("先自行作答、預測答案，或列出已知與未知的概念，保留自主認知投入。"),
                                createCell("『你原本怎麼想？』", true, BG_WARM)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("第二步", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                createCell("2. 再問", true),
                                createCell("帶著明確學習目的向 AI 提問，要求引導或提供線索，不直接索取最終成品。"),
                                createCell("『你這次提問想要解決什麼？』", false)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("第三步", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                createCell("3. 查證", true),
                                createCell("對照課本內文、實驗數據、官方資料庫或可靠第三方來源，檢驗 AI 正誤。"),
                                createCell("『AI 的哪句話有證據？』", true, BG_WARM)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("第四步", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                createCell("4. 取捨", true),
                                createCell("明確說明採納哪些內容、拒絕哪些幻覺或偏誤，以及自主修正的具體理由。"),
                                createCell("『你改了什麼？為什麼？』", true, BG_WARM)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("第五步", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                createCell("5. 揭露", true),
                                createCell("清楚標註 AI 協作範圍與工具名稱，並妥善保存對話與修改歷程紀錄。"),
                                createCell("『請標記哪部分來自 AI 協作。』", false)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("第六步", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                createCell("6. 脫離", true),
                                createCell("移開 AI 工具，能不看 AI 獨立口頭重述核心重點、用筆手寫或遷移至新情境題目。"),
                                createCell("『現在不看 AI，再說一次！』", true, BG_WARM)
                            ]
                        })
                    ]
                }),

                new Paragraph({ spacing: { after: 300 }, children: [] }),

                // Section 8: 七項學習設計指標
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    spacing: { before: 400, after: 180 },
                    children: [
                        new TextRun({
                            text: "伍、七項學習設計指標與課堂實踐檢核",
                            bold: true,
                            size: 32,
                            color: PRIMARY_COLOR
                        })
                    ]
                }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: [
                        new TableRow({
                            children: [
                                createHeaderCell("項次", 10),
                                createHeaderCell("觀察指標", 25),
                                createHeaderCell("可觀察證據 / 追問反思指引", 65)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("1", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                createCell("目標一致", true),
                                createCell("AI 活動是否直接服務領域學習目標？若移除 AI，核心學習目標仍然清楚完整嗎？")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("2", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                createCell("認知責任", true),
                                createCell("學生是否先想、作選擇、說理由，而非跳過大腦直接取得 AI 的完成品？")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("3", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                createCell("查證品質", true),
                                createCell("是否安排並落實使用課本、數據、實驗或可靠來源檢核 AI 生成內容？")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("4", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                createCell("人類主體性", true),
                                createCell("學生能否拒絕、修正 AI 建議，並維持自己的獨特觀點與原創創作意圖？")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("5", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                createCell("差異化支持", true),
                                createCell("針對不同學習基礎、語言能力、數位設備與自律程度的學生，是否有層次不同的支援鷹架？")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("6", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                createCell("獨立學習", true),
                                createCell("課堂是否確實安排「無 AI 口說、紙筆、實作或新情境任務」以檢驗知識是否內化遷移？")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("7", true, BG_LIGHT_GRAY, AlignmentType.CENTER),
                                createCell("安全與責任", true),
                                createCell("個人資料隱私、年齡規範、著作權法規、協作揭露、AI 偏誤與不當內容是否被妥善處置？")
                            ]
                        })
                    ]
                }),

                new Paragraph({ spacing: { after: 150 }, children: [] }),
                createCallout(
                    "教師教學觀察與反思句型公式",
                    "『我看到學生……，這可能支持/阻礙……；下一輪可嘗試……，再用……證據確認。』",
                    BG_LIGHT_BLUE,
                    PRIMARY_COLOR
                ),

                new Paragraph({ spacing: { after: 300 }, children: [] }),

                // Section 9: Bloom-AI 框架與不能外包的學習能力
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    spacing: { before: 400, after: 180 },
                    children: [
                        new TextRun({
                            text: "陸、Bloom-AI 框架與「不能讓 AI 外包」的學習能力",
                            bold: true,
                            size: 32,
                            color: PRIMARY_COLOR
                        })
                    ]
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "引據 Schober (2025) 的 BLOOM-AI Framework，AI 擅長支持基礎知識（AI supports foundational knowledge），而人類教師的關鍵使命在於培育高階認知技能（Human educators develop higher-order cognitive skills）。以下為各層級不能外包的核心要求：",
                            size: 22
                        })
                    ]
                }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: [
                        new TableRow({
                            children: [
                                createHeaderCell("Bloom 認知層次", 22),
                                createHeaderCell("學生應具備的獨立核心能力", 38),
                                createHeaderCell("【絕對不能讓 AI 外包】的防範準則", 40)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("創造 (Create)", true, BG_LIGHT_GRAY),
                                createCell("設定目的、重組內容，產出有意義的作品，並能說明創作決策。"),
                                createCell("不是只全盤接受 AI 生成的完整成品！必須有自主設計與決策理由。", true)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("評估 (Evaluate)", true, BG_LIGHT_GRAY),
                                createCell("比較多種解釋，查證來源與證據，做出有充分理由的判斷。"),
                                createCell("不是把 AI 提供的答案當作最後權威判斷！必須查證並說明取捨理由。", true)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("分析 (Analyze)", true, BG_LIGHT_GRAY),
                                createCell("整合多個線索，辨識因果、對比與論證關係，提出有證據的推論。"),
                                createCell("不是直接採用 AI 提出的推論！必須親自引用文本或數據證據進行推演。", true)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("應用 (Apply)", true, BG_LIGHT_GRAY),
                                createCell("提取規則 ➔ 辨識情境 ➔ 套用與轉換，在新情境中獨立選用規則。"),
                                createCell("不是照著 AI 範例生搬硬套模仿！必須在新情境中自主遷移驗證。", true)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("理解 (Understand)", true, BG_LIGHT_GRAY),
                                createCell("選擇重要線索，辨別重要與次要資訊，能用自己的話解釋。"),
                                createCell("不是複製貼上 AI 的摘要！必須用自己的詞彙與邏輯重新轉譯表達。", true)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("記憶 (Remember)", true, BG_LIGHT_GRAY),
                                createCell("從大腦記憶提取，或從選項中辨認並確認正確資訊。"),
                                createCell("不是全靠 AI 提示才想得起來！沒有提示時仍能獨立提取核心先備知識。", true)
                            ]
                        })
                    ]
                }),

                new Paragraph({ spacing: { after: 200 }, children: [] }),

                // Section 10: 國語文實務案例：理解題作答指導策略
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    spacing: { before: 300, after: 140 },
                    children: [
                        new TextRun({
                            text: "實務教學案例：國語文領域學力測驗理解題指導策略",
                            bold: true,
                            size: 26,
                            color: SECONDARY_COLOR
                        })
                    ]
                }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: [
                        new TableRow({
                            children: [
                                createHeaderCell("指導策略步驟", 25),
                                createHeaderCell("教師教學指導重點 (Instruction)", 38),
                                createHeaderCell("學生內化思考歷程 (Thinking Process)", 37)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("Step 1 找線索\n(Locate)", true, BG_LIGHT_GRAY),
                                createCell("• 回到文章中，找到相關句子或段落\n• 標記關鍵詞（因為、所以、其實、但是、這、他）"),
                                createCell("• 看題目問什麼意思\n• 回到文章尋找相關的句段並標註線索")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("Step 2 轉換說法\n(Paraphrase)", true, BG_LIGHT_GRAY),
                                createCell("• 把原文的說法「換句話說」\n• 避開照抄原句，改用同義詞或簡化句型"),
                                createCell("• 不照抄原句\n• 用自己的話把句意清楚說一次")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("Step 3 刪除次要資訊\n(Eliminate)", true, BG_LIGHT_GRAY),
                                createCell("• 使用嚴謹排除法\n• 過濾干擾選項與偏離主題的次要細節"),
                                createCell("• 刪掉完全不可能、不合邏輯的選項")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("Step 4 比對選項\n(Match)", true, BG_LIGHT_GRAY),
                                createCell("• 將「自己理解的意思」與選項逐一精準比對\n• 辨析最切合文本意涵的最佳解"),
                                createCell("• 找出與自己換句話說最接近、最合理的選項")
                            ]
                        })
                    ]
                }),

                new Paragraph({ spacing: { after: 200 }, children: [] }),

                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    spacing: { before: 300, after: 140 },
                    children: [
                        new TextRun({
                            text: "學生能力與國語文學科內容 Bloom 矩陣對照表",
                            bold: true,
                            size: 26,
                            color: SECONDARY_COLOR
                        })
                    ]
                }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: [
                        new TableRow({
                            children: [
                                createHeaderCell("Bloom 層次", 20),
                                createHeaderCell("學生應具備能力", 40),
                                createHeaderCell("對應之國語文學科學習內容", 40)
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("創造 (Create)", true, BG_LIGHT_GRAY),
                                createCell("設定目的、整合資訊，產出有意義的作品，並說明創作選擇"),
                                createCell("設計題目、改寫文本、重組觀點，形成有證據支持的作品")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("評估 (Evaluate)", true, BG_LIGHT_GRAY),
                                createCell("比較多種解釋，查證來源與證據，做出有理由的判斷"),
                                createCell("答案與選項品質、證據可信度、觀點評估、AI 回應查證")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("分析 (Analyze)", true, BG_LIGHT_GRAY),
                                createCell("整合多個線索，辨識因果、對比與論證關係，提出有證據解釋"),
                                createCell("推論、文章結構、作者觀點、論點與論據、修辭作用")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("應用 (Apply)", true, BG_LIGHT_GRAY),
                                createCell("提取規則 ➔ 辨識情境 ➔ 套用與轉換"),
                                createCell("句型、複句、修辭、詞語關係與文體規則的新情境運用、改寫與分類")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("理解 (Understand)", true, BG_LIGHT_GRAY),
                                createCell("用自己的話解釋，辨別重要與次要資訊"),
                                createCell("段落大意、指涉、因果、人物心情、舉例目的")
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("記憶 (Remember)", true, BG_LIGHT_GRAY),
                                createCell("從記憶提取，或從選項中辨識正確資訊"),
                                createCell("字音、字形、部首、詞義、筆順、詞類與規則")
                            ]
                        })
                    ]
                }),

                new Paragraph({ spacing: { after: 300 }, children: [] }),

                // Document Metadata / Conclusion
                new Paragraph({
                    spacing: { before: 200 },
                    children: [
                        new TextRun({
                            text: "───────────────────────────────────────────────────────",
                            color: "CBD5E0"
                        })
                    ]
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "本文件由 AI 視覺辨識模型從南投縣史港國小施君潔主任授課講義《AI導入課程與評量》完整轉譯、結構化編修而成，提供教育現場教師備課、研習、教案設計與數位教學評量之實務參考。",
                            size: 18,
                            italic: true,
                            color: "718096"
                        })
                    ]
                })
            ]
        }]
    });

    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync("AI導入課程與評量_實務架構手冊_施君潔.docx", buffer);
    console.log("Successfully generated AI導入課程與評量_實務架構手冊_施君潔.docx (Size: " + buffer.length + " bytes)");
}

buildDocument().catch(console.error);
