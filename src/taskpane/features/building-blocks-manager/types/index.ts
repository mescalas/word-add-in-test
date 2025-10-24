export interface ContentControlInfo {
  id: number;
  title: string;
  tag: string;
  type: string;
  appearance: string;
  cannotDelete: boolean;
  cannotEdit: boolean;
  color: string;
  placeholderText: string;
  text: string;
}

export interface CreateContentControlParams {
  title: string;
  tag: string;
  type: Word.ContentControlType;
  appearance?: Word.ContentControlAppearance;
  color?: string;
  placeholderText?: string;
  cannotDelete?: boolean;
  cannotEdit?: boolean;
}

export const CONTENT_CONTROL_TYPES = [
  { value: "RichText", label: "Rich Text" },
  { value: "PlainText", label: "Plain Text" },
  { value: "Paragraph", label: "Paragraph" },
  { value: "Picture", label: "Picture" },
  { value: "CheckBox", label: "Check Box" },
  { value: "ComboBox", label: "Combo Box" },
  { value: "DropDownList", label: "Drop Down List" },
  { value: "DatePicker", label: "Date Picker" },
  { value: "BuildingBlockGallery", label: "Building Block Gallery" },
] as const;
