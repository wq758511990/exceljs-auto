interface ExportColumnExtraProps {
  startIndex: number;
  endIndex: number;
  depth: number;
  childrenLength: number;
}

// 导入column的额外字段
interface ImportColumnExtraProps {
  validate: (val: unknown) => boolean;
  message: string;
  required: boolean;
  importHandler: (val: unknown) => any;
}

// 导入、导出数据类型集合
export interface TableColumnAllProps
  extends Partial<ExportColumnExtraProps>,
    Partial<ImportColumnExtraProps> {
  parent?: TableColumnAllProps;
  title?: unknown;
  dataIndex?: unknown;
  key?: unknown;
  width?: unknown;
  fixed?: unknown;
  render?: unknown;
  children?: TableColumnAllProps[];
}
