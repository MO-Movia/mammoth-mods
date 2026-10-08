export interface Mammoth {
    convertToHtml: (input: Input, options?: Options) => Promise<Result>;
    convert: (input: Input, options?: Options) => Promise<Result>;
}

export type Input = NodeJsInput | BrowserInput;

export type NodeJsInput = PathInput | BufferInput;

export interface PathInput {
    path: string;
}

export interface BufferInput {
    buffer: Buffer;
}

export type BrowserInput = ArrayBufferInput;

export interface ArrayBufferInput {
    arrayBuffer: ArrayBuffer;
}

export interface Options {
    styleMap?: string | Array<string>;
    includeEmbeddedStyleMap?: boolean;
    includeDefaultStyleMap?: boolean;
    convertImage?: ImageConverter;
    ignoreEmptyParagraphs?: boolean;
    idPrefix?: string;
    externalFileAccess?: boolean;
    transformDocument?: (element: any) => any;
}

export interface ImageConverter {
    __mammothBrand: "ImageConverter";
}

export interface Image {
    contentType: string;
    readAsArrayBuffer: () => Promise<ArrayBuffer>;
    readAsBase64String: () => Promise<string>;
    readAsBuffer: () => Promise<Buffer>;
    read: ImageRead;
}

export interface ImageRead {
    (): Promise<Buffer>;
    (encoding: string): Promise<string>;
}

export interface ImageAttributes {
    src: string;
}

export interface Images {
    dataUri: ImageConverter;
    imgElement: (f: (image: Image) => Promise<ImageAttributes>) => ImageConverter;
}

export interface Result {
    value: string;
    messages: Array<Message>;
}

export type Message = Warning | Error;

export interface Warning {
    type: "warning";
    message: string;
}

export interface Error {
    type: "error";
    message: string;
    error: unknown;
}

export declare function convertToHtml(input: Input, options?: Options): Promise<Result>;

export declare function convert(input: Input, options?: Options): Promise<Result>;

declare const mammoth: Mammoth;

export default mammoth;
