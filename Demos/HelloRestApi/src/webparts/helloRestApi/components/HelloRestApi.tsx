import * as React from 'react';
import styles from './HelloRestApi.module.scss';
import type { IHelloRestApiProps } from './IHelloRestApiProps';
import { ISPHttpClientOptions, SPHttpClient } from '@microsoft/sp-http';

export default class HelloRestApi extends React.Component<IHelloRestApiProps> {
  private doGetText = async (restUrl: string): Promise<string> => {
    let result = "";

    const response = await this.props.spHttpClient.get(restUrl, SPHttpClient.configurations.v1);
    result = `Status: ${response.status}\n`;
    if (response.ok) {
      const data = await response.json();
      result += JSON.stringify(data, null, 2);
    } else {
      result += await response.text();
    }

    return result;
  }

  private doGetJson = async (restUrl: string): Promise<any> => {
    let result = null;

    const response = await this.props.spHttpClient.get(restUrl, SPHttpClient.configurations.v1);
    if (response.ok) {
      const data = await response.json();
      result = data;
    } else {
      const message = await response.text();
      throw new Error(message);
    }

    return result;
  }

  private doPost = async (restUrl: string, body?: any, headers?: any): Promise<string> => {
    let result = "";

    try {
      let options: ISPHttpClientOptions = {
        body: body,
        headers: headers
      };

      const response = await this.props.spHttpClient.post(restUrl, 
        SPHttpClient.configurations.v1, options);
      result = `Status: ${response.status}\n`;
      if (response.ok) {
        const data = await response.json();
        result += JSON.stringify(data, null, 2);
      } else {
        result += await response.text();
      }
    } catch (ex) {
      console.log("error", ex);
    }

    return result;
  }

  private doMerge = async (restUrl: string, body: any, etag: string = "*"): Promise<string> => {
    const headers = {
      'X-HTTP-Method': 'MERGE',
      'IF-MATCH': etag
    }

    return await this.doPost(restUrl, body, headers);
  }

  private doDelete = async (restUrl: string, etag: string = "*"): Promise<string> => {
    const headers = {
      'X-HTTP-Method': 'DELETE',
      'IF-MATCH': etag
    }

    return await this.doPost(restUrl, undefined, headers);
  }

  private getFileBuffer = async (file: File): Promise<ArrayBuffer> => {
    return new Promise<ArrayBuffer>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e: any) => {
        resolve(e.target.result);
      };
      reader.onerror = (e: any) => {
        reject(e.target.error);
      };
      reader.readAsArrayBuffer(file);
    });
  }

  private doUploadFile = async (restUrl: string, file: File): Promise<string> => {
    const buffer = await this.getFileBuffer(file);

    const headers = {
      'CONTENT-LENGTH': buffer.byteLength.toString()
    }

    return await this.doPost(restUrl, buffer, headers);
  }

  private getWebButtonClick = async () => {
    const restUrl = this.props.webAbsoluteUrl + "/_api/web";

    const output = this.props.domElement.querySelector('#output') as HTMLTextAreaElement;
    output.value = await this.doGetText(restUrl);    
  };

  private getListsButtonClick = async () => {
    const restUrl = this.props.webAbsoluteUrl + "/_api/web/lists";

    const output = this.props.domElement.querySelector('#output') as HTMLTextAreaElement;
    output.value = await this.doGetText(restUrl);    
  };

  private getListsWithSelectButtonClick = async () => {
    const restUrl = this.props.webAbsoluteUrl + 
      "/_api/web/lists?$select=Title,ItemCount";

    const output = this.props.domElement.querySelector('#output') as HTMLTextAreaElement;
    output.value = await this.doGetText(restUrl);    
  };

  private getListsWithFilterButtonClick = async () => {
    const restUrl = this.props.webAbsoluteUrl + 
      "/_api/web/lists?$select=Title,ItemCount&" + 
      "$filter=((Hidden eq false) and (ItemCount gt 0))";

    const output = this.props.domElement.querySelector('#output') as HTMLTextAreaElement;
    output.value = await this.doGetText(restUrl);    
  };

  private getWithExpand1ButtonClick = async () => {
    const restUrl = this.props.webAbsoluteUrl + 
      "/_api/web/?$select=Title,Lists/Title,Lists/Hidden,Lists/ItemCount&$expand=Lists";

    const output = this.props.domElement.querySelector('#output') as HTMLTextAreaElement;
    output.value = await this.doGetText(restUrl);    
  };

  private getWithExpand2ButtonClick = async () => {
    const restUrl = this.props.webAbsoluteUrl + 
      "/_api/Web/Lists/getByTitle('Products')/Items?" + 
      "$select=Title,Category/Title&" + 
      "$filter=(Category/Title eq 'Beverages')&" + 
      "$expand=Category/Title";

    const output = this.props.domElement.querySelector('#output') as HTMLTextAreaElement;
    output.value = await this.doGetText(restUrl);    
  };

  private createTasksListButtonClick = async () => {
    const restUrl = this.props.webAbsoluteUrl + "/_api/web/lists";
    const body = JSON.stringify({
      BaseTemplate: 107,
      Title: "Tasks"    
    });

    const output = this.props.domElement.querySelector('#output') as HTMLTextAreaElement;
    output.value = await this.doPost(restUrl, body);    
  };

  private createTasksListItemButtonClick = async () => {
    const restUrl = this.props.webAbsoluteUrl + 
      "/_api/web/lists/getByTitle('Tasks')/items";
    const dueDate = new Date();
    dueDate.setDate(dueDate.getDate() + 7);
    const body = JSON.stringify({
      Title: "Sample Task",
      AssignedToId: this.props.legacyPageContext.userId,
      DueDate: dueDate
    });

    const output = this.props.domElement.querySelector('#output') as HTMLTextAreaElement;
    output.value = await this.doPost(restUrl, body);    
  };

  private updateTasksListItemButtonClick = async () => {
    let restUrl = this.props.webAbsoluteUrl + 
      "/_api/web/lists/getByTitle('Tasks')/items?$top=1";
    const data = await this.doGetJson(restUrl);

    const items = data.value;
    if (items.length === 1) {
      const item = items[0];

      restUrl = this.props.webAbsoluteUrl + 
        `/_api/web/lists/getByTitle('Tasks')/items(${item.Id})`;
      const body = JSON.stringify({
        Status: "In Progress",
        PercentComplete: 0.10
      });

      const output = this.props.domElement.querySelector('#output') as HTMLTextAreaElement;
      output.value = await this.doMerge(restUrl, body, item["@odata.etag"]);    
    }
  };

  private deleteTasksListItemButtonClick = async () => {
    let restUrl = this.props.webAbsoluteUrl + 
      "/_api/web/lists/getByTitle('Tasks')/items?$top=1";
    const data = await this.doGetJson(restUrl);

    const items = data.value;
    if (items.length === 1) {
      const item = items[0];

      restUrl = this.props.webAbsoluteUrl + 
        `/_api/web/lists/getByTitle('Tasks')/items(${item.Id})`;

      const output = this.props.domElement.querySelector('#output') as HTMLTextAreaElement;
      output.value = await this.doDelete(restUrl, item["@odata.etag"]);    
    }
  };

  private uploadFileButtonClick = async () => {
    const output = this.props.domElement.querySelector('#output') as HTMLTextAreaElement;

    const fileInput = this.props.domElement.querySelector('#uploadFile') as HTMLInputElement;
    if (fileInput.files!.length !== 1) {
      output.value = "Please select a file to upload.";
      return;
    }

    const file = fileInput.files![0];
    const restUrl = this.props.webAbsoluteUrl + 
      `/_api/web/lists/GetByTitle('Documents')/RootFolder/Files/add(` +
      `overwrite=true, url='${file.name}')`;
    
    output.value = await this.doUploadFile(restUrl, file);
  };

  public render(): React.ReactElement<IHelloRestApiProps> {
    return (
      <section className={`${styles.helloRestApi} ${this.props.hasTeamsContext ? styles.teams : ''}`}>
        <table>
          <tr>
            <td style={{verticalAlign: "top"}}>
              <input type="button" value="Get Web" onClick={this.getWebButtonClick}  /><br />
              <input type="button" value="Get Lists" onClick={this.getListsButtonClick}  /><br />
              <input type="button" value="Get Lists with Select" onClick={this.getListsWithSelectButtonClick} /><br />
              <input type="button" value="Get Lists with Filter" onClick={this.getListsWithFilterButtonClick} /><br />
              <input type="button" value="Get with Expand 1" onClick={this.getWithExpand1ButtonClick} /><br />
              <input type="button" value="Get with Expand 2" onClick={this.getWithExpand2ButtonClick} /><br />
              <input type="button" value="Create Tasks List" onClick={this.createTasksListButtonClick} /><br />
              <input type="button" value="Create Tasks List Item" onClick={this.createTasksListItemButtonClick} /><br />              
              <input type="button" value="Update Tasks List Item" onClick={this.updateTasksListItemButtonClick} /><br />
              <input type="button" value="Delete Tasks List Item" onClick={this.deleteTasksListItemButtonClick} /><br />
              <input type="button" value="Upload File" onClick={this.uploadFileButtonClick} /><br />
              <input type="file" id="uploadFile"/><br />
            </td>
            <td style={{width: "20px"}}>
                &nbsp;
            </td>
            <td style={{verticalAlign: "top"}}>
              <textarea id="output" rows={25} cols={80}>
                &nbsp;
              </textarea>
            </td>
          </tr>
        </table>
      </section>
    );
  }
}
