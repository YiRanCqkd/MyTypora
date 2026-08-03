export interface MarkdownSourceToken {
  type: string;
  map?: [number, number];
  info?: string;
  content?: string;
  attrSet?: (name: string, value: string) => void;
  attrGet?: (name: string) => string | null;
}
