export interface WordApiService {
  initialize(): Promise<void>;
  getSelectedText(): Promise<string>;
  getEntireDocumentText(): Promise<string>;
  replaceSelectedText(text: string): Promise<void>;
}

class WordApiImpl implements WordApiService {

  async initialize(): Promise<void> {
    return new Promise((resolve, reject) => {
      if (typeof Office === 'undefined') {
        reject(new Error('Office.js not loaded'));
        return;
      }

      Office.onReady((info) => {
        if (info.host === Office.HostType.Word) {
          resolve();
        } else {
          reject(new Error('Not running in Word'));
        }
      });
    });
  }


  async getSelectedText(): Promise<string> {
    return new Promise((resolve, reject) => {
      Word.run(async (context) => {
        try {
          const selection = context.document.getSelection();
          selection.load('text');
          await context.sync();
          resolve(selection.text);
        } catch (error) {
          reject(error);
        }
      });
    });
  }

  async getEntireDocumentText(): Promise<string> {
    return new Promise((resolve, reject) => {
      Word.run(async (context) => {
        try {
          const body = context.document.body;
          body.load('text');
          await context.sync();
          resolve(body.text);
        } catch (error) {
          reject(error);
        }
      });
    });
  }

  async replaceSelectedText(text: string): Promise<void> {
    return new Promise((resolve, reject) => {
      Word.run(async (context) => {
        try {
          const selection = context.document.getSelection();
          selection.insertText(text, Word.InsertLocation.replace);
          await context.sync();
          resolve();
        } catch (error) {
          reject(error);
        }
      });
    });
  }

}

export const wordApi: WordApiService = new WordApiImpl();