import { WorkbookProject } from './types';

export const EMPTY_WORKBOOK: WorkbookProject = {
  fileName: 'Untitled_Workbook.xlsx',
  lastModified: new Date().toISOString(),
  sheets: [
    {
      id: 'sheet_1',
      name: 'Sheet 1',
      description: '',
      hiddenState: 0,
      headerRowIndex: -1,
      fields: [
        {
          id: 'sheet_1_c0_col_a',
          key: 'col_a',
          label: 'A',
          sourceHeader: '(No Header — Col A)',
          columnLetter: 'A',
          type: 'text',
          visible: true,
          width: 160,
          defaultValue: '',
          validation: {},
        },
        {
          id: 'sheet_1_c1_col_b',
          key: 'col_b',
          label: 'B',
          sourceHeader: '(No Header — Col B)',
          columnLetter: 'B',
          type: 'text',
          visible: true,
          width: 160,
          defaultValue: '',
          validation: {},
        },
        {
          id: 'sheet_1_c2_col_c',
          key: 'col_c',
          label: 'C',
          sourceHeader: '(No Header — Col C)',
          columnLetter: 'C',
          type: 'text',
          visible: true,
          width: 160,
          defaultValue: '',
          validation: {},
        },
      ],
      rows: [
        {
          id: 'sheet_1_r0',
          excelRowNumber: 1,
          cells: {
            sheet_1_c0_col_a: '',
            sheet_1_c1_col_b: '',
            sheet_1_c2_col_c: '',
          },
        },
      ],
    },
  ],
};
