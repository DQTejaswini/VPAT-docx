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
                          size: 22,
                          font: "Arial",
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
                          size: 22,
                          font: "Arial",
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
                          size: 22,
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
              text: "EN 301 549 Report",
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
              text: "Chapter 4: ",
              bold: true,
              size: 32,
              font: "Arial",
              color: "auto",
            }),
            new TextRun({
              text: "Functional Performance Statements (FPS)",
              style: "Hyperlink",
              size: 32,
              bold: true,
              font: "Arial",
              color: "0000FF",
              underline: {},
              hyperlink:
                "https://www.etsi.org/deliver/etsi_en/301500_301599/301549/03.02.01_60/en_301549v030201p.pdf#page=20",
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
                  width: { size: 40, type: WidthType.PERCENTAGE },
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
                }),
                new TableCell({
                  width: { size: 20, type: WidthType.PERCENTAGE },
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
                }),
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
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
                          text: "4.2.1 Usage without vision",
                          bold: true,
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
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
                    ...[
                      "1.1.1 Non-Text Content",
                      "1.3.1 Info and Relationships",
                      "1.3.2 Meaningful Sequence",
                      "1.3.3 Sensory Characteristics",
                      "1.4.1 Use of Color",
                      "2.1.1 Keyboard",
                      "2.1.2 No Keyboard Trap",
                      "2.4.2 Page Titled",
                      "2.4.3 Focus Order",
                      "2.4.4 Link Purpose (in context)",
                      "2.4.6 Headings and Labels",
                      "3.1.1 Language of Page",
                      "3.3.1 Error Identification",
                      "4.1.2 Name, Role, Value",
                      "4.1.3 Status Messages",
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
                          text: "4.2.2 Usage with limited vision",
                          bold: true,
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
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
                          text: "4.2.3 Usage without perception of colour",
                          bold: true,
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
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
                          text: "4.2.4 Usage without hearing",
                          bold: true,
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
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
                          text: "4.2.5 Usage with limited hearing",
                          bold: true,
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
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
                          text: "4.2.6 Usage with no or limited vocal capability",
                          bold: true,
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
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
                          text: "4.2.7 Usage with limited manipulation or strength",
                          bold: true,
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
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
                          text: "4.2.8 Usage with limited reach",
                          bold: true,
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
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
                    ...[
                      "2.5.3 Label in Name",
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
                          text: "4.2.9 Minimize photosensitive seizure triggers",
                          bold: true,
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
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
                }),
              ],
            }),

            //Data Row 10
            new TableRow({
              children: [
                // Criteria
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "4.2.10 Usage with limited cognition, language or learning",
                          bold: true,
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
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
                }),

                // Remarks and Explanations
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Web: Most functionality is usable by people with limited language, cognitive, and learning abilities. People with cognitive disabilities have varying needs for features that allow them to adapt content and work with assistive technology. Exceptions are noted in:",
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                    ...[
                      "1.1.1 Non-Text Content",
                      "1.3.1 Info and Relationships",
                      "1.3.2 Meaningful Sequence",
                      "1.3.3 Sensory Characteristics",
                      "1.4.1 Use of Color",
                      "1.4.10 Non-Text Contrast",
                      "1.4.12 Text Spacing",
                      "2.1.1 Keyboard",
                      "2.4.2 Page Titled",
                      "2.4.3 Focus Order",
                      "2.4.4 Link Purpose (in context)",
                      "2.4.6 Headings and Labels",
                      "2.4.7 Focus Visible",
                      "2.5.3 Label in Name",
                      "3.1.1 Language of Page",
                      "3.3.1 Error Identification",
                      "3.3.2 Labels or Instructions",
                      "4.1.2 Name, Role, Value",
                      "4.1.3 Status Messages",
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
                }),
              ],
            }),

            //Data Row 11
            new TableRow({
              children: [
                // Criteria
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "4.2.11 Privacy",
                          bold: true,
                          font: "Arial",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
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
                }),

                // Remarks and Explanations
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Web: Where the product provides features for accessibility, it maintains the privacy of people who use these features at the same level as other users.",
                          font: "Arial",
                          size: 22,
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
