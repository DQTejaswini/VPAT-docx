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
            }),
          ],
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
            }),
          ],
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
            }),
          ],
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
            }),
          ],
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
            }),
          ],
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
            }),
          ],
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
                          size: 24,
                          font: "Calibri",
                          color: "auto",
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  shading: {
                    fill: "D9D9D9",
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "URL",
                          bold: true,
                          size: 24,
                          font: "Calibri",
                          color: "auto",
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  shading: {
                    fill: "D9D9D9",
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
                          size: 24,
                          font: "Calibri",
                          color: "auto",
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
                          text: "https://patient.health-ce.wolterskluwer.com/demo_login/login",
                          size: 24,
                          font: "Calibri",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
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
                          size: 24,
                          font: "Calibri",
                          color: "auto",
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
                          text: "https://patient.health-ce.wolterskluwer.com/demo_team/search",
                          size: 24,
                          font: "Calibri",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
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
                          size: 24,
                          font: "Calibri",
                          color: "auto",
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
                          text: "https://patient.health-ce.wolterskluwer.com/results/Top%20Surgery",
                          size: 24,
                          font: "Calibri",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
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
                          size: 24,
                          font: "Calibri",
                          color: "auto",
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
                          text: "https://patient.health-ce.wolterskluwer.com/multimedia/686679F8-2E8A-405C-A1B2-2C956A14267E/43688",
                          size: 24,
                          font: "Calibri",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
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
                          size: 24,
                          font: "Calibri",
                          color: "auto",
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
                          text: "https://patient.health-ce.wolterskluwer.com/article/686679F8-2E8A-405C-A1B2-2C956A14267E/44750",
                          size: 24,
                          font: "Calibri",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
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
                          size: 24,
                          font: "Calibri",
                          color: "auto",
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
                          text: "https://patient.health-ce.wolterskluwer.com/demo_results/programs",
                          size: 24,
                          font: "Calibri",
                          color: "0563C1",
                        }),
                      ],
                    }),
                  ],
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
                          font: "Calibri",
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
                          size: 24,
                          font: "Calibri",
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  shading: { fill: "BFBFBF" },
                }),
                new TableCell({
                  width: { size: 20, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Conformance Level",
                          bold: true,
                          size: 24,
                          font: "Calibri",
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  shading: { fill: "BFBFBF" },
                }),
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Remarks and Explanations",
                          bold: true,
                          size: 24,
                          font: "Calibri",
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  shading: { fill: "BFBFBF" },
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
                          font: "Calibri",
                          size: 24,
                          color: "0563C1",
                          underline: {},
                          hyperlink:
                            "https://www.w3.org/WAI/WCAG21/quickref/#non-text-content",
                        }),
                        new TextRun({
                          text: " (Level A)",
                          font: "Calibri",
                          size: 24,
                        }),
                      ],
                    }),
                  ],
                }),
                new TableCell({
                  width: { size: 20, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Partially Supports",
                          size: 24,
                          font: "Calibri",
                        }),
                      ],
                    }),
                  ],
                }),
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Most non-text content has text alternatives or a text alternate that serves an equivalent purpose. The following exception exists:",
                          size: 24,
                          font: "Calibri",
                        }),
                      ],
                    }),
                    new Paragraph({
                      bullet: { level: 0 },
                      children: [
                        new TextRun({
                          text: "A complex image does not have a long description to convey the information presented by the image, so people who are blind and/or use a screen reader will not be able to understand the information presented by the image. This occurs on the following page: Heart Healthy Diet - Article.",
                          size: 24,
                          font: "Calibri",
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("accessibility_report.docx", buffer);
  console.log("Document created successfully!");
});
