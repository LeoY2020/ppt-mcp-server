declare module 'unzipper' {
  import { Transform } from 'stream';

  interface Entry {
    path: string;
    type: 'File' | 'Directory';
    buffer(): Promise<Buffer>;
    autodrain(): void;
    pipe<T extends NodeJS.WritableStream>(destination: T): T;
  }

  interface ParseStream extends Transform {
    on(event: 'entry', listener: (entry: Entry) => void): this;
    on(event: 'close', listener: () => void): this;
    on(event: 'error', listener: (error: Error) => void): this;
  }

  export function Parse(): ParseStream;
}
