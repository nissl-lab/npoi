import { Tutorial } from './types';

export const TUTORIALS: Tutorial[] = [
  {
    id: 1,
    title: "Create a simple Excel file with NPOI",
    description: "Learn the basics of NPOI by creating your first Excel workbook, adding a sheet, and writing data to cells.",
    videoUrl: "https://www.youtube.com/watch?v=PUCdied11dM&list=PL7J6yRMWV1ot32hsgCZJ5sI4Fp2QKovCM&index=1",
    category: 'Excel',
    difficulty: 'Beginner'
  },
  {
    id: 2,
    title: "Open an existing Excel template and setup formula",
    description: "Discover how to load an existing .xlsx file, modify its content, and programmatically set up Excel formulas.",
    videoUrl: "https://www.youtube.com/watch?v=kPNaeS_TVNg&list=PL7J6yRMWV1ot32hsgCZJ5sI4Fp2QKovCM&index=9",
    category: 'Excel',
    difficulty: 'Intermediate'
  },
  {
    id: 3,
    title: "Applying data formats to Excel cells",
    description: "Master cell styling by applying custom data formats for dates, currency, and numbers using NPOI styles.",
    videoUrl: "https://www.youtube.com/watch?v=w2aEKIuns0E&list=PL7J6yRMWV1ot32hsgCZJ5sI4Fp2QKovCM&index=8",
    category: 'Excel',
    difficulty: 'Beginner'
  },
  {
    id: 4,
    title: "Insert pictures in Excel with NPOI",
    description: "Enhance your spreadsheets by learning how to embed images and position them precisely within your Excel sheets.",
    videoUrl: "https://www.youtube.com/watch?v=KO3TuOi_Smc&list=PL7J6yRMWV1ot32hsgCZJ5sI4Fp2QKovCM&index=7",
    category: 'Excel',
    difficulty: 'Intermediate'
  },
  {
    id: 5,
    title: "How to use WorkbookFactory",
    description: "Understand the WorkbookFactory pattern for creating and opening workbooks in a version-agnostic way.",
    videoUrl: "https://www.youtube.com/watch?v=H4FzMWkZNH8&list=PL7J6yRMWV1ot32hsgCZJ5sI4Fp2QKovCM&index=6",
    category: 'General',
    difficulty: 'Beginner'
  },
  {
    id: 6,
    title: "How to operate Sheet object",
    description: "Deep dive into the ISheet interface: renaming sheets, setting active sheets, and managing sheet properties.",
    videoUrl: "https://www.youtube.com/watch?v=Njurg1aEPTY&list=PL7J6yRMWV1ot32hsgCZJ5sI4Fp2QKovCM&index=5",
    category: 'Excel',
    difficulty: 'Beginner'
  },
  {
    id: 7,
    title: "How to set fill pattern color for a cell",
    description: "Learn how to use ICellStyle to set background colors and fill patterns for cells to make your data pop.",
    videoUrl: "https://www.youtube.com/watch?v=K5X-TkBLN_s&list=PL7J6yRMWV1ot32hsgCZJ5sI4Fp2QKovCM&index=4",
    category: 'Excel',
    difficulty: 'Intermediate'
  },
  {
    id: 8,
    title: "Create a Word document with Paragraph and Run",
    description: "Transition to Word processing with NPOI. Learn about XWPFDocument, Paragraphs, and Text Runs.",
    videoUrl: "https://www.youtube.com/watch?v=I_hoiZrsjMI&list=PL7J6yRMWV1ot32hsgCZJ5sI4Fp2QKovCM&index=3",
    category: 'Word',
    difficulty: 'Beginner'
  },
  {
    id: 9,
    title: "Create a hyperlink in a Word document",
    description: "Add interactivity to your Word documents by programmatically inserting external and internal hyperlinks.",
    videoUrl: "https://www.youtube.com/watch?v=ah0n3jiFnlg&list=PL7J6yRMWV1ot32hsgCZJ5sI4Fp2QKovCM&index=2",
    category: 'Word',
    difficulty: 'Intermediate'
  },
  {
    id: 11,
    title: "Add a new Line in Word paragraph",
    description: "Control text flow in Word by learning how to insert line breaks and manage paragraph spacing.",
    videoUrl: "https://www.youtube.com/watch?v=xFFcn_g5rmc&list=PL7J6yRMWV1ot32hsgCZJ5sI4Fp2QKovCM&index=1",
    category: 'Word',
    difficulty: 'Beginner'
  },
  {
    id: 15,
    title: "How to insert a picture in a Word document",
    description: "Learn the specific APIs for embedding images into XWPF Word documents using NPOI.",
    videoUrl: "https://www.youtube.com/watch?v=B0zh6H8W3EE&list=PL7J6yRMWV1ot32hsgCZJ5sI4Fp2QKovCM",
    category: 'Word',
    difficulty: 'Intermediate'
  },
  {
    id: 16,
    title: "Use ExcelMapper to generate an Excel file",
    description: "Explore ExcelMapper, a powerful extension for NPOI that simplifies mapping objects to Excel rows.",
    videoUrl: "https://www.youtube.com/watch?v=YdqDeoVqGlo&list=PL7J6yRMWV1ot32hsgCZJ5sI4Fp2QKovCM&index=1",
    category: 'Excel',
    difficulty: 'Advanced'
  }
];
