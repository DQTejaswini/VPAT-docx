const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  SectionType,
  PageOrientation,
  AlignmentType,
  Header,
  Footer,
  ImageRun,
  PageNumber,
  NumberOfTotalPages,
  HeadingLevel,
  Table,
  TableRow,
  TableCell,
  WidthType,
  Bullet,
  BulletType,
} = require("docx");
const fs = require("fs");
const path = require("path");

const base64Image = (path) => {
  try {
    const bitmap = fs.readFileSync(path);
    return new Buffer(bitmap).toString("base64");
  } catch (error) {
    console.error("Error reading image:", error);
    return null;
  }
};

const imagePath = "/Users/dq_tejaswini/Desktop/VPAT automation/deque.png";
const imageBuffer = base64Image(imagePath);

const doc = new Document({
  sections: [
    {
      properties: {
        type: SectionType.NEXT_PAGE,
        page: {
          orientation: PageOrientation.LANDSCAPE,
          size: { width: 16838, height: 11906 },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              children: [
                imageBuffer
                  ? new ImageRun({
                      data: imageBuffer,
                      transformation: {
                        width: 165,
                        height: 70,
                      },
                      altText: {
                        name: "Deque Logo",
                        description: "Deque Logo",
                      },
                    })
                  : new TextRun("Image not found"),
              ],
              alignment: AlignmentType.RIGHT,
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "________________________________________________",
                  bold: true,
                  size: 12,
                  font: "Arial",
                }),
              ],
              alignment: AlignmentType.LEFT,
              spacing: {
                before: 100,
                after: 100,
              },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: '"Voluntary Product Accessibility Template" and "VPAT" are registered',
                  size: 22,
                  font: "Calibri",
                }),
                new TextRun({
                  text: "service marks of the Information Technology Industry Council (ITI)",
                  size: 22,
                  font: "Calibri",
                  break: 1,
                }),
              ],
              alignment: AlignmentType.LEFT,
              spacing: {
                before: 100,
                after: 100,
              },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "Page ",
                  size: 22,
                  font: "Calibri",
                }),
                new TextRun({
                  children: [PageNumber.CURRENT],
                  size: 22,
                  font: "Calibri",
                }),
                new TextRun({
                  text: " of ",
                  size: 22,
                  font: "Calibri",
                }),
                new TextRun({
                  children: [PageNumber.TOTAL_PAGES],
                  size: 22,
                  font: "Calibri",
                }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: {
                before: 100,
              },
            }),
          ],
        }),
      },
      children: [
        new Paragraph({
          children: [
            new TextRun({
              text: "Wolters Kluwer – SSW Accessibility Conformance Report",
              bold: true,
              alignment: "CENTER",
              size: 48,
              font: "Arial",
              color: "auto",
            }),
          ],
          heading: HeadingLevel.HEADING_1,
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "(Based on VPAT® Version 2.5)",
              bold: true,
              size: 24,
              font: "Arial",
            }),
          ],
          alignment: AlignmentType.CENTER,
          spacing: {
            before: 300,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Name of Product/Version:",
              bold: true,
              size: 36,
              font: "Arial",
              color: "auto",
            }),
          ],
          heading: HeadingLevel.HEADING_2,
          spacing: {
            before: 300,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Wolters Kluwer – SSW",
              size: 24,
              font: "Arial",
            }),
          ],
          spacing: {
            before: 150,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Report Date:",
              bold: true,
              size: 36,
              font: "Arial",
              color: "auto",
            }),
          ],
          heading: HeadingLevel.HEADING_2,
          spacing: {
            before: 300,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "14 July 2025",
              size: 24,
              font: "Arial",
            }),
          ],
          spacing: {
            before: 150,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Product Description:",
              bold: true,
              size: 36,
              font: "Arial",
              color: "auto",
            }),
          ],
          heading: HeadingLevel.HEADING_2,
          spacing: {
            before: 300,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "TBD",
              size: 24,
              font: "Arial",
            }),
          ],
          spacing: {
            before: 150,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Contact Information:",
              bold: true,
              size: 36,
              font: "Arial",
              color: "auto",
            }),
          ],
          heading: HeadingLevel.HEADING_2,
          spacing: {
            before: 300,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "TBD",
              size: 24,
              font: "Arial",
            }),
          ],
          spacing: {
            before: 150,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Notes:",
              bold: true,
              size: 36,
              font: "Arial",
              color: "auto",
            }),
          ],
          heading: HeadingLevel.HEADING_2,
          spacing: {
            before: 300,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "This report was created by Deque Systems Inc. upon completion of an accessibility evaluation performed between 21 November, 2024 and 27 November, 2024. The validation of issues was performed in the month of June 2025",
              size: 24,
              font: "Arial",
            }),
          ],
          spacing: {
            before: 150,
          },
        }),
      ],
    },
    {
      properties: {
        type: SectionType.NEXT_PAGE,
        page: {
          orientation: PageOrientation.LANDSCAPE,
          size: { width: 16838, height: 11906 },
          margin: {
            top: 300,
            right: 720,
            bottom: 720,
            left: 720,
          },
          headerDistance: 284,
        },
      },
      headers: {
        default: new Header({
          children: [],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "Page ",
                  size: 22,
                  font: "Calibri",
                }),
                new TextRun({
                  children: [PageNumber.CURRENT],
                  size: 22,
                  font: "Calibri",
                }),
                new TextRun({
                  text: " of ",
                  size: 22,
                  font: "Calibri",
                }),
                new TextRun({
                  children: [PageNumber.TOTAL_PAGES],
                  size: 22,
                  font: "Calibri",
                }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: {
                before: 100,
              },
            }),
          ],
        }),
      },
      children: [
        new Paragraph({
          children: [
            new TextRun({
              text: "Evaluation Methods Used:",
              bold: true,
              size: 36,
              font: "Arial",
              color: "auto",
            }),
          ],
          heading: HeadingLevel.HEADING_2,
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 300,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "A combination of automated and manual testing techniques was employed for the accessibility assessment. Manual assessment was performed using Chrome on Windows and included exclusive use of the keyboard. Automated tools used included axe Auditor and the axe Dev Tools browser extension. Assistive technologies employed included latest version of NVDA.",
              size: 24,
              font: "Arial",
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 150,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Scope of Evaluation",
              bold: true,
              size: 36,
              font: "Arial",
              color: "auto",
            }),
          ],
          heading: HeadingLevel.HEADING_2,
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 300,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "The pages in the following table were evaluated as part of the assessment on which this report is based.",
              size: 24,
              font: "Arial",
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 150,
            after: 300,
          },
        }),
        new Table({
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Page Title",
                          bold: true,
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  shading: {
                    fill: "D9D9D9",
                  },
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "URL",
                          bold: true,
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  shading: {
                    fill: "D9D9D9",
                  },
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Login",
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "https://patient.health-ce.wolterskluwer.com/demo_login/login",
                          size: 22,
                          font: "Arial",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Search",
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "https://patient.health-ce.wolterskluwer.com/demo_team/search",
                          size: 22,
                          font: "Arial",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Search Results",
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "https://patient.health-ce.wolterskluwer.com/results/Top%20Surgery",
                          size: 22,
                          font: "Arial",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Top Surgery - Interactive Video",
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "https://patient.health-ce.wolterskluwer.com/multimedia/686679F8-2E8A-405C-A1B2-2C956A14267E/43688",
                          size: 22,
                          font: "Arial",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Heart Healthy Diet - Article",
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "https://patient.health-ce.wolterskluwer.com/article/686679F8-2E8A-405C-A1B2-2C956A14267E/44750",
                          size: 22,
                          font: "Arial",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Program List (With Search Disabled)",
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "https://patient.health-ce.wolterskluwer.com/demo_results/programs",
                          size: 22,
                          font: "Arial",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
          ],
          width: {
            size: 100,
            type: WidthType.PERCENTAGE,
          },
        }),

        new Paragraph({
          spacing: {
            before: 400,
            after: 400,
          },
          children: [
            new TextRun({
              text: "In addition to the pages listed above, the following components that appear on multiple pages were tested as part of the assessment:",
              size: 24,
              font: "Arial",
            }),
          ],
          alignment: AlignmentType.LEFT,
        }),
        new Paragraph({
          alignment: AlignmentType.LEFT,
          bullet: {
            level: 0,
          },
          children: [
            new TextRun({
              text: "Header before login",
              size: 24,
              bold: false,
              font: "Arial",
              color: "#000000",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.LEFT,
          bullet: {
            level: 0,
          },
          children: [
            new TextRun({
              text: "Header after login",
              size: 24,
              bold: false,
              font: "Arial",
              color: "#000000",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.LEFT,
          bullet: {
            level: 0,
          },
          children: [
            new TextRun({
              text: "Footer",
              size: 24,
              bold: false,
              font: "Arial",
              color: "#000000",
            }),
          ],
        }),
      ],
    },
    {
      properties: {
        type: SectionType.NEXT_PAGE,
        page: {
          orientation: PageOrientation.LANDSCAPE,
          size: { width: 16838, height: 11906 },
          margin: {
            top: 300,
            right: 720,
            bottom: 720,
            left: 720,
          },
          headerDistance: 284,
        },
      },
      headers: {
        default: new Header({
          children: [],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "Page ",
                  size: 22,
                  font: "Calibri",
                }),
                new TextRun({
                  children: [PageNumber.CURRENT],
                  size: 22,
                  font: "Calibri",
                }),
                new TextRun({
                  text: " of ",
                  size: 22,
                  font: "Calibri",
                }),
                new TextRun({
                  children: [PageNumber.TOTAL_PAGES],
                  size: 22,
                  font: "Calibri",
                }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: {
                before: 100,
              },
            }),
          ],
        }),
      },
      children: [
        new Paragraph({
          children: [
            new TextRun({
              text: "Evaluation Methods Used:",
              bold: true,
              size: 36,
              font: "Arial",
              color: "auto",
            }),
          ],
          heading: HeadingLevel.HEADING_2,
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 300,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "A combination of automated and manual testing techniques was employed for the accessibility assessment. Manual assessment was performed using Chrome on Windows and included exclusive use of the keyboard. Automated tools used included axe Auditor and the axe Dev Tools browser extension. Assistive technologies employed included latest version of NVDA.",
              size: 24,
              font: "Arial",
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 150,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Scope of Evaluation",
              bold: true,
              size: 36,
              font: "Arial",
              color: "auto",
            }),
          ],
          heading: HeadingLevel.HEADING_2,
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 300,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "The pages in the following table were evaluated as part of the assessment on which this report is based.",
              size: 24,
              font: "Arial",
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 150,
            after: 300,
          },
        }),
        new Table({
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Page Title",
                          bold: true,
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  shading: {
                    fill: "D9D9D9",
                  },
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "URL",
                          bold: true,
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  shading: {
                    fill: "D9D9D9",
                  },
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Login",
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "https://patient.health-ce.wolterskluwer.com/demo_login/login",
                          size: 22,
                          font: "Arial",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Search",
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "https://patient.health-ce.wolterskluwer.com/demo_team/search",
                          size: 22,
                          font: "Arial",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Search Results",
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "https://patient.health-ce.wolterskluwer.com/results/Top%20Surgery",
                          size: 22,
                          font: "Arial",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Top Surgery - Interactive Video",
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "https://patient.health-ce.wolterskluwer.com/multimedia/686679F8-2E8A-405C-A1B2-2C956A14267E/43688",
                          size: 22,
                          font: "Arial",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Heart Healthy Diet - Article",
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "https://patient.health-ce.wolterskluwer.com/article/686679F8-2E8A-405C-A1B2-2C956A14267E/44750",
                          size: 22,
                          font: "Arial",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Program List (With Search Disabled)",
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "https://patient.health-ce.wolterskluwer.com/demo_results/programs",
                          size: 22,
                          font: "Arial",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
          ],
          width: {
            size: 100,
            type: WidthType.PERCENTAGE,
          },
        }),

        new Paragraph({
          spacing: {
            before: 400,
            after: 400,
          },
          children: [
            new TextRun({
              text: "In addition to the pages listed above, the following components that appear on multiple pages were tested as part of the assessment:",
              size: 24,
              font: "Arial",
            }),
          ],
          alignment: AlignmentType.LEFT,
        }),
        new Paragraph({
          alignment: AlignmentType.LEFT,
          bullet: {
            level: 0,
          },
          children: [
            new TextRun({
              text: "Header before login",
              size: 24,
              bold: false,
              font: "Arial",
              color: "#000000",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.LEFT,
          bullet: {
            level: 0,
          },
          children: [
            new TextRun({
              text: "Header after login",
              size: 24,
              bold: false,
              font: "Arial",
              color: "#000000",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.LEFT,
          bullet: {
            level: 0,
          },
          children: [
            new TextRun({
              text: "Footer",
              size: 24,
              bold: false,
              font: "Arial",
              color: "#000000",
            }),
          ],
        }),
      ],
    },
    {
      properties: {
        type: SectionType.NEXT_PAGE,
        page: {
          orientation: PageOrientation.LANDSCAPE,
          size: { width: 16838, height: 11906 },
          margin: {
            top: 300,
            right: 720,
            bottom: 720,
            left: 720,
          },
          headerDistance: 284,
        },
      },
      headers: {
        default: new Header({
          children: [],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "Page ",
                  size: 22,
                  font: "Calibri",
                }),
                new TextRun({
                  children: [PageNumber.CURRENT],
                  size: 22,
                  font: "Calibri",
                }),
                new TextRun({
                  text: " of ",
                  size: 22,
                  font: "Calibri",
                }),
                new TextRun({
                  children: [PageNumber.TOTAL_PAGES],
                  size: 22,
                  font: "Calibri",
                }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: {
                before: 100,
              },
            }),
          ],
        }),
      },
      children: [
        new Paragraph({
          children: [
            new TextRun({
              text: "Evaluation Methods Used:",
              bold: true,
              size: 36,
              font: "Arial",
              color: "auto",
            }),
          ],
          heading: HeadingLevel.HEADING_2,
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 300,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "A combination of automated and manual testing techniques was employed for the accessibility assessment. Manual assessment was performed using Chrome on Windows and included exclusive use of the keyboard. Automated tools used included axe Auditor and the axe Dev Tools browser extension. Assistive technologies employed included latest version of NVDA.",
              size: 24,
              font: "Arial",
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 150,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Scope of Evaluation",
              bold: true,
              size: 36,
              font: "Arial",
              color: "auto",
            }),
          ],
          heading: HeadingLevel.HEADING_2,
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 300,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "The pages in the following table were evaluated as part of the assessment on which this report is based.",
              size: 24,
              font: "Arial",
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 150,
            after: 300,
          },
        }),
        new Table({
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Page Title",
                          bold: true,
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  shading: {
                    fill: "D9D9D9",
                  },
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "URL",
                          bold: true,
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  shading: {
                    fill: "D9D9D9",
                  },
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Login",
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "https://patient.health-ce.wolterskluwer.com/demo_login/login",
                          size: 22,
                          font: "Arial",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Search",
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "https://patient.health-ce.wolterskluwer.com/demo_team/search",
                          size: 22,
                          font: "Arial",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Search Results",
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "https://patient.health-ce.wolterskluwer.com/results/Top%20Surgery",
                          size: 22,
                          font: "Arial",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Top Surgery - Interactive Video",
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "https://patient.health-ce.wolterskluwer.com/multimedia/686679F8-2E8A-405C-A1B2-2C956A14267E/43688",
                          size: 22,
                          font: "Arial",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Heart Healthy Diet - Article",
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "https://patient.health-ce.wolterskluwer.com/article/686679F8-2E8A-405C-A1B2-2C956A14267E/44750",
                          size: 22,
                          font: "Arial",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Program List (With Search Disabled)",
                          size: 22,
                          font: "Arial",
                          color: "auto",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "https://patient.health-ce.wolterskluwer.com/demo_results/programs",
                          size: 22,
                          font: "Arial",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
          ],
          width: {
            size: 100,
            type: WidthType.PERCENTAGE,
          },
        }),

        new Paragraph({
          spacing: {
            before: 400,
            after: 400,
          },
          children: [
            new TextRun({
              text: "In addition to the pages listed above, the following components that appear on multiple pages were tested as part of the assessment:",
              size: 24,
              font: "Arial",
            }),
          ],
          alignment: AlignmentType.LEFT,
        }),
        new Paragraph({
          alignment: AlignmentType.LEFT,
          bullet: {
            level: 0,
          },
          children: [
            new TextRun({
              text: "Header before login",
              size: 24,
              bold: false,
              font: "Arial",
              color: "#000000",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.LEFT,
          bullet: {
            level: 0,
          },
          children: [
            new TextRun({
              text: "Header after login",
              size: 24,
              bold: false,
              font: "Arial",
              color: "#000000",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.LEFT,
          bullet: {
            level: 0,
          },
          children: [
            new TextRun({
              text: "Footer",
              size: 24,
              bold: false,
              font: "Arial",
              color: "#000000",
            }),
          ],
        }),
      ],
    },
    {
      properties: {
        type: SectionType.NEXT_PAGE,
        page: {
          orientation: PageOrientation.LANDSCAPE,
          size: { width: 16838, height: 11906 },
          margin: {
            top: 300,
            right: 720,
            bottom: 720,
            left: 720,
          },
          headerDistance: 284,
        },
      },
      headers: {
        default: new Header({
          children: [],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "Page ",
                  size: 22,
                  font: "Calibri",
                }),
                new TextRun({
                  children: [PageNumber.CURRENT],
                  size: 22,
                  font: "Calibri",
                }),
                new TextRun({
                  text: " of ",
                  size: 22,
                  font: "Calibri",
                }),
                new TextRun({
                  children: [PageNumber.TOTAL_PAGES],
                  size: 22,
                  font: "Calibri",
                }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: {
                before: 100,
              },
            }),
          ],
        }),
      },
      children: [
        new Paragraph({
          children: [
            new TextRun({
              text: "Applicable Standards/Guidelines",
              bold: true,
              size: 36,
              font: "Arial",
              color: "auto",
            }),
          ],
          heading: HeadingLevel.HEADING_2,
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 300,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "This report covers the degree of conformance for the following accessibility standard/guidelines:",
              size: 24,
              font: "Arial",
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 150,
            after: 300,
          },
        }),

        new Table({
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Standard/Guideline",
                          bold: true,
                          size: 24,
                          font: "Arial",
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  shading: { fill: "BFBFBF" },
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Included In Report",
                          bold: true,
                          size: 24,
                          font: "Arial",
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  shading: { fill: "BFBFBF" },
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),

            // Row for WCAG 2.0
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Web Content Accessibility Guidelines 2.0",
                          style: "Hyperlink",
                          size: 22,
                          font: "Arial",
                          color: "0000FF",
                          underline: {},
                          hyperlink: "https://www.w3.org/WAI/WCAG20/quickref/",
                        }),
                      ],
                    }),
                  ],
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Level A (Yes)",
                          size: 22,
                          font: "Arial",
                          alignment: AlignmentType.CENTER,
                        }),
                      ],
                    }),
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Level AA (Yes)",
                          size: 22,
                          font: "Arial",
                          alignment: AlignmentType.CENTER,
                        }),
                      ],
                    }),
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Level AAA (No)",
                          size: 22,
                          font: "Arial",
                          alignment: AlignmentType.CENTER,
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),

            // Row for WCAG 2.1
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Web Content Accessibility Guidelines 2.1",
                          style: "Hyperlink",
                          size: 22,
                          font: "Arial",
                          color: "0000FF",
                          underline: {},
                          hyperlink: "https://www.w3.org/WAI/WCAG21/quickref/",
                        }),
                      ],
                    }),
                  ],
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Level A (Yes)",
                          size: 22,
                          font: "Arial",
                          alignment: AlignmentType.CENTER,
                        }),
                      ],
                    }),
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Level AA (Yes)",
                          size: 22,
                          font: "Arial",
                          alignment: AlignmentType.CENTER,
                        }),
                      ],
                    }),
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Level AAA (No)",
                          size: 22,
                          font: "Arial",
                          alignment: AlignmentType.CENTER,
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),

            // Row for WCAG 2.2
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Web Content Accessibility Guidelines 2.2",
                          style: "Hyperlink",
                          size: 22,
                          font: "Arial",
                          color: "0000FF",
                          underline: {},
                          hyperlink: "https://www.w3.org/WAI/WCAG22/quickref/",
                        }),
                      ],
                    }),
                  ],
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Level A (Yes)",
                          size: 22,
                          font: "Arial",
                          alignment: AlignmentType.CENTER,
                        }),
                      ],
                    }),
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Level AA (Yes)",
                          size: 22,
                          font: "Arial",
                          alignment: AlignmentType.CENTER,
                        }),
                      ],
                    }),
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Level AAA (No)",
                          size: 22,
                          font: "Arial",
                          alignment: AlignmentType.CENTER,
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          ],
          width: {
            size: 80,
            type: WidthType.PERCENTAGE,
            alignment: AlignmentType.CENTER,
          },
        }),

        new Paragraph({
          children: [
            new TextRun({
              text: "Terms",
              bold: true,
              size: 36,
              font: "Arial",
              color: "auto",
            }),
          ],
          heading: HeadingLevel.HEADING_2,
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 300,
          },
        }),

        new Paragraph({
          children: [
            new TextRun({
              text: "The terms used in the Conformance Level information are defined as follows:",
              size: 24,
              font: "Arial",
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 150,
            after: 300,
          },
        }),

        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 250,
          },
          bullet: {
            level: 0,
          },
          children: [
            new TextRun({
              text: "Supports: ",
              size: 24,
              bold: true,
              font: "Arial",
              color: "#000000",
            }),
            new TextRun({
              text: "The functionality of the product has at least one method that meets the criterion without known defects or meets with equivalent facilitation.",
              size: 24,
              bold: false,
              font: "Arial",
            }),
          ],
        }),

        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 150,
          },
          bullet: {
            level: 0,
          },
          children: [
            new TextRun({
              text: "Partially Supports: ",
              size: 24,
              bold: true,
              font: "Arial",
              color: "#000000",
            }),
            new TextRun({
              text: "Some functionality of the product does not meet the criterion.",
              size: 24,
              bold: false,
              font: "Arial",
            }),
          ],
        }),

        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 150,
          },
          bullet: {
            level: 0,
          },
          children: [
            new TextRun({
              text: "Does Not Support: ",
              size: 24,
              bold: true,
              font: "Arial",
              color: "#000000",
            }),
            new TextRun({
              text: "The majority of product functionality does not meet the criterion.",
              size: 24,
              bold: false,
              font: "Arial",
            }),
          ],
        }),

        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 150,
          },
          bullet: {
            level: 0,
          },
          children: [
            new TextRun({
              text: "Not Applicable: ",
              size: 24,
              bold: true,
              font: "Arial",
              color: "#000000",
            }),
            new TextRun({
              text: "The criterion is not relevant to the product.",
              size: 24,
              bold: false,
              font: "Arial",
            }),
          ],
        }),

        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 150,
          },
          bullet: {
            level: 0,
          },
          children: [
            new TextRun({
              text: "Not Evaluated: ",
              size: 24,
              bold: true,
              font: "Arial",
              color: "#000000",
            }),
            new TextRun({
              text: "The product has not been evaluated against the criterion. This can be used only in WCAG Level AAA criteria.",
              size: 24,
              bold: false,
              font: "Arial",
            }),
          ],
        }),

        new Paragraph({
          children: [
            new TextRun({
              text: "WCAG 2.2 Report",
              bold: true,
              size: 36,
              font: "Arial",
              color: "auto",
            }),
          ],
          heading: HeadingLevel.HEADING_2,
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 300,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Note: When reporting on conformance with the WCAG 2.2 Success Criteria, the criteria are scoped for full pages, complete processes, and accessibility-supported ways of using technology as documented in the ",
              size: 24,
              font: "Arial",
              color: "auto",
            }),
            new TextRun({
              text: "WCAG 2.2 Conformance Requirements.",
              style: "Hyperlink",
              size: 24,
              font: "Arial",
              color: "0000FF",
              underline: {},
              hyperlink: "https://www.w3.org/TR/WCAG22/#conformance-reqs",
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 300,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Table 1: Success Criteria, Level A",
              bold: true,
              size: 32,
              font: "Arial",
              color: "auto",
            }),
          ],
          heading: HeadingLevel.HEADING_3,
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 300,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Notes:",
              size: 24,
              font: "Arial",
              color: "auto",
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 150,
          },
        }),
        new Table({
          alignment: AlignmentType.CENTER,
          width: {
            size: 100,
            type: WidthType.PERCENTAGE,
          },
          rows: [
            // Header Row
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Criteria",
                          bold: true,
                          size: 22,
                          font: "Arial",
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  shading: { fill: "BFBFBF" },
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 20, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Conformance Level",
                          bold: true,
                          size: 22,
                          font: "Arial",
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  shading: { fill: "BFBFBF" },
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Remarks and Explanations",
                          bold: true,
                          size: 22,
                          font: "Arial",
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  shading: { fill: "BFBFBF" },
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),

            // First Data Row (1.1.1 Non-text Content)
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "1.1.1 Non-text Content",
                          font: "Arial",
                          size: 22,
                          color: "0563C1",
                          underline: {},
                          hyperlink:
                            "https://www.w3.org/WAI/WCAG21/quickref/#non-text-content",
                        }),
                        new TextRun({
                          text: " (Level A)",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 20, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Partially Supports",
                          size: 22,
                          font: "Arial",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Most non-text content has text alternatives or a text alternate that serves an equivalent purpose. The following exception exists:",
                          size: 22,
                          font: "Arial",
                        }),
                      ],
                    }),
                    new Paragraph({
                      children: [],
                      spacing: { before: 100 }
                    }),
                    new Paragraph({
                      bullet: { level: 0 },
                      children: [
                        new TextRun({
                          text: "A complex image does not have a long description to convey the information presented by the image, so people who are blind and/or use a screen reader will not be able to understand the information presented by the image. This occurs on the following page: Heart Healthy Diet - Article.",
                          size: 22,
                          font: "Arial",
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
          ],
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Table 3: Success Criteria, Level AAA",
              bold: true,
              size: 32,
              font: "Arial",
              color: "auto",
            }),
          ],
          heading: HeadingLevel.HEADING_3,
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 300,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Notes: Not Applicable. The Universal Widgets Website was not assessed for WCAG 2.1 Level AAA conformance.",
              size: 24,
              font: "Arial",
              color: "auto",
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 150,
          },
        }),
      ],
    },
    {
      properties: {
        type: SectionType.NEXT_PAGE,
        page: {
          orientation: PageOrientation.LANDSCAPE,
          size: { width: 16838, height: 11906 },
          margin: {
            top: 300,
            right: 720,
            bottom: 720,
            left: 720,
          },
          headerDistance: 284,
        },
      },
      headers: {
        default: new Header({
          children: [],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "Page ",
                  size: 22,
                  font: "Calibri",
                }),
                new TextRun({
                  children: [PageNumber.CURRENT],
                  size: 22,
                  font: "Calibri",
                }),
                new TextRun({
                  text: " of ",
                  size: 22,
                  font: "Calibri",
                }),
                new TextRun({
                  children: [PageNumber.TOTAL_PAGES],
                  size: 22,
                  font: "Calibri",
                }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: {
                before: 100,
              },
            }),
          ],
        }),
      },
      children: [
        new Paragraph({
          children: [
            new TextRun({
              text: "Revised Section 508 Report",
              bold: true,
              size: 36,
              font: "Arial",
              color: "auto",
            }),
          ],
          heading: HeadingLevel.HEADING_2,
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 300,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Notes:",
              size: 24,
              font: "Arial",
              color: "auto",
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 150,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Chapter 3: ",
              bold: true,
              size: 32,
              font: "Arial",
              color: "auto",
            }),
            new TextRun({
              text: "Functional Performance Criteria (FPC)",
              style: "Hyperlink",
              size: 32,
              bold: true,
              font: "Arial",
              color: "0000FF",
              underline: {},
              hyperlink:
                "https://www.access-board.gov/ict/#chapter-3-functional-performance-criteria",
            }),
          ],
          heading: HeadingLevel.HEADING_3,
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 250,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Notes:",
              size: 24,
              font: "Arial",
              color: "auto",
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 150,
          },
        }),
        new Table({
          alignment: AlignmentType.CENTER,
          width: {
            size: 100,
            type: WidthType.PERCENTAGE,
          },
          rows: [
            // Header Row
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 46.1, type: WidthType.PERCENTAGE },
                  shading: { fill: "BFBFBF" },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Criteria",
                          bold: true,
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 27, type: WidthType.PERCENTAGE },
                  shading: { fill: "BFBFBF" },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Conformance Level",
                          bold: true,
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 26.9, type: WidthType.PERCENTAGE },
                  shading: { fill: "BFBFBF" },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Remarks and Explanations",
                          bold: true,
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),

            //Data Row 1
            new TableRow({
              children: [
                // Criteria
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "302.1 Without vision",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),

                // Conformance Level
                new TableCell({
                  width: { size: 20, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Web: Does Not Support",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),

                // Remarks and Explanations
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Web: Most, if not all, functionality is not usable without vision. Examples are noted in:",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                    new Paragraph({
                      children: [],
                      spacing: { before: 100 }
                    }),
                    ...[
                      "1.1.1 Non-Text Content",
                      "1.3.1 Info and Relationships",
                      "1.4.1 Use of Color",
                      "2.1.1 Keyboard",
                      "2.4.2 Page Titled",
                      "2.4.4 Link Purpose (in context)",
                      "2.4.6 Headings and Labels",
                      "3.1.2 Language of Parts",
                      "3.3.1 Error Identification",
                      "4.1.2 Name, Role, Value",
                    ].map(
                      (item) =>
                        new Paragraph({
                          bullet: { level: 0 },
                          children: [
                            new TextRun({
                              text: item,
                              font: "Arial",
                              size: 22,
                            }),
                          ],
                        })
                    ),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),

            //Data Row 2
            new TableRow({
              children: [
                // Criteria
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "302.2 With limited vision",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),

                // Conformance Level
                new TableCell({
                  width: { size: 20, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Web: Does Not Support",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),

                // Remarks and Explanations
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Web: Most, if not all, functionality is not usable with limited vision. Examples are noted in: ",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                    new Paragraph({
                      children: [],
                      spacing: { before: 100 }
                    }),
                    ...[
                      "1.1.1 Non-Text Content",
                      "1.3.1 Info and Relationships",
                      "1.3.2 Meaningful Sequence",
                      "1.3.3 Sensory Characteristics",
                      "1.4.1 Use of Color",
                      "1.4.3 Contrast (minimum)",
                      "1.4.4 Resize Text",
                      "1.4.10 Non-Text Contrast",
                      "1.4.11 Non-Text Contrast",
                      "1.4.12 Text Spacing",
                      "2.1.1 Keyboard",
                      "2.1.2 No Keyboard Trap",
                      "2.4.2 Page Titled",
                      "2.4.3 Focus Order",
                      "2.4.4 Link Purpose (in context)",
                      "2.4.6 Headings and Labels",
                      "2.4.7 Focus Visible",
                      "3.1.1 Language of Page",
                      "3.3.1 Error Identification",
                      "3.3.2 Labels or Instructions",
                      "4.1.2 Name, Role, Value",
                    ].map(
                      (item) =>
                        new Paragraph({
                          bullet: { level: 0 },
                          children: [
                            new TextRun({
                              text: item,
                              font: "Arial",
                              size: 22,
                            }),
                          ],
                        })
                    ),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),

            //Data Row 3
            new TableRow({
              children: [
                // Criteria
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "302.3 Without Perception of Color",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),

                // Conformance Level
                new TableCell({
                  width: { size: 20, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Web: Does Not Support",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),

                // Remarks and Explanations
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Web: Most, if not all, functionality is not usable without perception of color. Examples are noted in: ",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                    new Paragraph({
                      children: [],
                      spacing: { before: 100 }
                    }),
                    ...[
                      "1.4.1 Use of Color",
                      "1.4.3 Contrast (minimum)",
                      "1.4.11 Non-Text Contrast",
                      "3.3.1 Error Identification",
                    ].map(
                      (item) =>
                        new Paragraph({
                          bullet: { level: 0 },
                          children: [
                            new TextRun({
                              text: item,
                              font: "Arial",
                              size: 22,
                            }),
                          ],
                        })
                    ),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),

            //Data Row 4
            new TableRow({
              children: [
                // Criteria
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "302.4 Without Hearing",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),

                // Conformance Level
                new TableCell({
                  width: { size: 20, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Web: Partially Supports",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),

                // Remarks and Explanations
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Web: Most functionality is usable without hearing. Exceptions are noted in:",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                    new Paragraph({
                      children: [],
                      spacing: { before: 100 }
                    }),
                    ...[
                      "1.1.1 Non-Text Content",
                      "1.3.3 Sensory Characteristics",
                      "3.1.1 Language of Page",
                    ].map(
                      (item) =>
                        new Paragraph({
                          bullet: { level: 0 },
                          children: [
                            new TextRun({
                              text: item,
                              font: "Arial",
                              size: 22,
                            }),
                          ],
                        })
                    ),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),

            //Data Row 5
            new TableRow({
              children: [
                // Criteria
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "302.5 With Limited Hearing",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),

                // Conformance Level
                new TableCell({
                  width: { size: 20, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Web: Partially Supports",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),

                // Remarks and Explanations
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Web: Most functionality is usable with limited hearing. Exceptions are noted in:",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                    new Paragraph({
                      children: [],
                      spacing: { before: 100 }
                    }),
                    ...[
                      "1.1.1 Non-Text Content",
                      "1.3.3 Sensory Characteristics",
                      "3.1.1 Language of Page",
                    ].map(
                      (item) =>
                        new Paragraph({
                          bullet: { level: 0 },
                          children: [
                            new TextRun({
                              text: item,
                              font: "Arial",
                              size: 22,
                            }),
                          ],
                        })
                    ),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),

            //Data Row 6
            new TableRow({
              children: [
                // Criteria
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "302.6 Without Speech",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),

                // Conformance Level
                new TableCell({
                  width: { size: 20, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Web: Not Applicable",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),

                // Remarks and Explanations
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Web: The product does not require the use of speech or other vocal output.",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),

            //Data Row 7
            new TableRow({
              children: [
                // Criteria
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "302.7 With Limited Manipulation",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),

                // Conformance Level
                new TableCell({
                  width: { size: 20, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Web: Does Not Support",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),

                // Remarks and Explanations
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Web: Most, if not all, functionality is not usable by people with limited manipulation and/or requires manipulation, simultaneous action, or hand strength. Examples are noted in:",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                    new Paragraph({
                      children: [],
                      spacing: { before: 100 }
                    }),
                    ...[
                      "2.1.1 Keyboard",
                      "2.1.2 No Keyboard Trap",
                      "2.4.2 Page Titled",
                      "2.4.3 Focus Order",
                      "2.4.4 Link Purpose (in context)",
                      "2.4.6 Headings and Labels",
                      "2.4.7 Focus Visible",
                      "2.5.3 Label in Name",
                      "4.1.2 Name, Role, Value",
                    ].map(
                      (item) =>
                        new Paragraph({
                          bullet: { level: 0 },
                          children: [
                            new TextRun({
                              text: item,
                              font: "Arial",
                              size: 22,
                            }),
                          ],
                        })
                    ),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),

            //Data Row 8
            new TableRow({
              children: [
                // Criteria
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "302.8 With Limited Reach and Strength",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),

                // Conformance Level
                new TableCell({
                  width: { size: 20, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Web: Supports",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),

                // Remarks and Explanations
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Web: All functionality is usable by people with limited reach.",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                    new Paragraph({
                      children: [],
                      spacing: { before: 100 }
                    }),
                    ...["2.5.3 Label in Name"].map(
                      (item) =>
                        new Paragraph({
                          bullet: { level: 0 },
                          children: [
                            new TextRun({
                              text: item,
                              font: "Arial",
                              size: 22,
                            }),
                          ],
                        })
                    ),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),

            //Data Row 9
            new TableRow({
              children: [
                // Criteria
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "302.9 With Limited Language, Cognitive, and Learning Abilities",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),

                // Conformance Level
                new TableCell({
                  width: { size: 20, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Web: Supports",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),

                // Remarks and Explanations
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Web: All functionality may be operated in a mode that minimizes the potential for triggering photosensitive seizures.",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
          ],
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Chapter 4: ",
              bold: true,
              size: 32,
              font: "Arial",
              color: "auto",
            }),
            new TextRun({
              text: "Hardware",
              style: "Hyperlink",
              size: 32,
              bold: true,
              font: "Arial",
              color: "0000FF",
              underline: {},
              hyperlink:
                "https://www.access-board.gov/ict/#chapter-5-software",
            }),
          ],
          heading: HeadingLevel.HEADING_3,
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 250,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Notes: The ICT covered by this report is not hardware. As such, the requirements of this chapter do not apply.",
              size: 24,
              font: "Arial",
              color: "auto",
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 150,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Chapter 5: ",
              bold: true,
              size: 32,
              font: "Arial",
              color: "auto",
            }),
            new TextRun({
              text: "Software",
              style: "Hyperlink",
              size: 32,
              bold: true,
              font: "Arial",
              color: "0000FF",
              underline: {},
              hyperlink:
                "https://www.access-board.gov/ict/#chapter-5-software",
            }),
          ],
          heading: HeadingLevel.HEADING_3,
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 250,
          },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Notes:",
              size: 24,
              font: "Arial",
              color: "auto",
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: {
            before: 150,
          },
        }),
        new Table({
          alignment: AlignmentType.CENTER,
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            // Header Row
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 46.1, type: WidthType.PERCENTAGE },
                  shading: { fill: "BFBFBF" },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Criteria",
                          bold: true,
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 27, type: WidthType.PERCENTAGE },
                  shading: { fill: "BFBFBF" },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Conformance Level",
                          bold: true,
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 26.9, type: WidthType.PERCENTAGE },
                  shading: { fill: "BFBFBF" },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Remarks and Explanations",
                          bold: true,
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),

            // First Data Row
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 46.1, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "501.1 Scope – Incorporation of WCAG 2.0 AA",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 27, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "See WCAG 2.1 section",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 29.7, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "See information in WCAG 2.1 section",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                  shading: { fill: "E7E6E6" },
                }),
              ],
            }),

            // Second Data Row
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "502 Interoperability with Assistive Technology",
                          bold: true,
                          italics: true,
                          font: "Arial",
                          size: 22,
                          color: "0000FF",
                          underline: {},
                          hyperlink: "https://www.access-board.gov/ict/#502-interoperability-assistive-technology"
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                  shading: { fill: "E7E6E6" },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Heading cell – no response required",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                  shading: { fill: "E7E6E6" },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Heading cell – no response required",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                  shading: { fill: "E7E6E6" },
                }),
              ],
            }),

            // Third Data Row
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "5.1.2.1 Closed functionality",
                          bold: true,
                          italics: true,
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "See 5.2 through 13",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "See information in 5.2 through 13",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),

            // Fourth Data Row
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "5.1.2.2 Assistive technology",
                          bold: true,
                          italics: true,
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "See 5.1.3 through 5.1.6",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "See information in 5.1.3 through 5.1.6",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),

            // Fifth Data Row
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "5.1.3 Non-visual access",
                          bold: true,
                          italics: true,
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Heading cell – no response required",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Heading cell – no response required",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),

            // Sixth Data Row
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "5.1.3.1 Audio output of visual information",
                          bold: true,
                          italics: true,
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Not Applicable",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "The website does not have closed functionality and allows the use of an outside or add-on screen reader to provide audio output of visual information.",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),

            // Sixth Data Row
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "5.1.3.2 Auditory output delivery including speech",
                          bold: true,
                          italics: true,
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Not Applicable",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "The website does not have closed functionality and allows auditory output (including speech) to be conveyed by outside mechanisms (including speakers, handsets, and earbuds).",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),

            // Seventh Data Row
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "5.1.3.3 Auditory output correlation",
                          bold: true,
                          italics: true,
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Not Applicable",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "The website does not have closed functionality, so audio information is conveyed using outside or add-on software and/or hardware.",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),

            // Eighth Data Row
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "5.1.3.4 Speech output user control",
                          bold: true,
                          italics: true,
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Not Applicable",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "The website does not have closed functionality and allows auditory output (including speech) to be conveyed by outside software (e.g., screen readers) and outside mechanisms (including speakers, handsets, and earbuds).",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),

            // Nineth Data Row
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "5.1.3.5 Speech output automatic interruption",
                          bold: true,
                          italics: true,
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Not Applicable",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "The website does not have closed functionality and allows auditory output (including speech) to be conveyed by outside software (e.g., screen readers) outside mechanisms (including speakers, handsets, and earbuds).",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),

            // Tenth Data Row
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "5.1.3.6 Speech output for non-text content",
                          bold: true,
                          italics: true,
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Not Applicable",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "The website does not have closed functionality and allows auditory output (including the announcement of text alternatives for images) to be conveyed by outside software (e.g., screen readers) and outside mechanisms (including speakers, handsets, and earbuds).",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),

            // Eleventh Data Row
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "5.1.3.7 Speech output for video information",
                          bold: true,
                          italics: true,
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Not Applicable",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "The website does not have closed functionality, and pre-recorded video content is not needed to enable the use of closed functions. Also, the product allows auditory output (including audio description of visual information in a video) to be conveyed by outside software (e.g., screen readers) and outside mechanisms (including speakers, handsets, and earbuds).",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),

            // Twelfth Data Row
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "5.1.3.8 Masked entry",
                          bold: true,
                          italics: true,
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Not Applicable",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "The website does not have closed functionality, so any auditory output of masked characters is controlled and conveyed by outside software (e.g., screen readers) and outside mechanisms (including speakers, handsets, and earbuds).",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),

            // Thirteenth Data Row
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "5.1.3.9 Private access to personal data",
                          bold: true,
                          italics: true,
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Not Applicable",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "The website does not have closed functionality that would provide private access to personal data.",
                          font: "Arial",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                  margins: {
                    left: 100,
                  },
                }),
              ],
            }),
          ],
        }),
      ],
    },
  ],
});

// Generate the document
Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("VPAT_Report.docx", buffer);
  console.log("Document created successfully!");
});