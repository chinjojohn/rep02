import * as ko from 'knockout';
import styles from './KnockoutWebpart.module.scss';
import { IKnockoutWebpartWebPartProps } from './KnockoutWebpartWebPart';
import { IWebPartContext } from '@microsoft/sp-webpart-base';  
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';  

export interface IKnockoutWebpartBindingContext extends IKnockoutWebpartWebPartProps {
  shouter: KnockoutSubscribable<{}>;
  context: IWebPartContext;

}

export interface IListItem {  
    Id: number;  
    Title: string;  
  }  
  

export default class KnockoutWebpartViewModel {
  public _koListName: KnockoutObservable<string> = ko.observable('');
   public message: KnockoutObservable<string>=ko.observable('');
   private _context: IWebPartContext;
   private _listName: string;
  

  public knockoutWebpartClass: string = styles.knockoutWebpart;
  public containerClass: string = styles.container;
  public rowClass: string = styles.row;
  public columnClass: string = styles.column;
  public titleClass: string = styles.title;
  public subTitleClass: string = styles.subTitle;
  public descriptionClass: string = styles.description;
  public buttonClass: string = styles.button;
  public labelClass: string = styles.label;

  constructor(bindings: IKnockoutWebpartBindingContext) {
    this._koListName(bindings.listName);
    this._context = bindings.context;
this._listName = bindings.listName;

    // When web part description is updated, change this view model's description.
    bindings.shouter.subscribe((value: string) => {
      this._koListName(value);
      this._listName=value;
    }, this, 'listName');
}
private getLatestItemId(): Promise<number> {  
      return this._context.spHttpClient.get(this._context.pageContext["web"]["absoluteUrl"]  
        + `/_api/web/lists/GetByTitle('${this._listName}')/items?$orderby=Id desc&$top=1&$select=id`, SPHttpClient.configurations.v1)  
        .then((response: SPHttpClientResponse): Promise<any> => {  
          return response.json();  
        })  
        .then((data: any): number => {  
          this.message("Load succeeded");  
          return data.value[0].ID;  
        },  
        (error: any) => {  
          this.message("Load failed");  
        }) as Promise<number>;  
    }
  private createItem(): void {
        const body: string = JSON.stringify({
          'Title': `Item ${new Date()}`
        });
        this._context.spHttpClient.post(`${this._context.pageContext["web"]["absoluteUrl"]}/_api/web/lists/getbytitle('${this._listName}')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {        'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          },
          body: body
        })
        .then((response: SPHttpClientResponse): Promise<IListItem> => {
          return response.json();
        })
        .then((item: IListItem): void => {
          this.message(`Item '${item.Title}' (ID: ${item.Id}) successfully created`);
        }, (error: any): void => {
          this.message('Error while creating the item: ' + error);
        });
      }
    private readItem(): void {  
          this.getLatestItemId()  
            .then((itemId: number): Promise<SPHttpClientResponse> => {  
              if (itemId === -1) {  
                throw new Error('No items found in the list');  
              }  
              this.message(`Loading information about item ID: ${itemId}...`);  
              return this._context.spHttpClient.get(`${this._context.pageContext["web"]["absoluteUrl"]}/_api/web/lists/getbytitle('${this._listName}')/items(${itemId})?$select=Title,Id`,  
                SPHttpClient.configurations.v1,  
                {  
                  headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'odata-version': ''
                  } 
                });  
            })  
            .then((response: SPHttpClientResponse): Promise<IListItem> => {  
              return response.json();  
            })  
            .then((item: IListItem): void => {  
              this.message(`Item ID: ${item.Id}, Title: ${item.Title}`);  
            }, (error: any): void => {  
              this.message('Loading latest item failed with error: ' + error);  
            });  
        }    
      private updateItem(): void {
            let latestItemId: number = undefined;
            this.message('Loading latest item...');
            this.getLatestItemId()
              .then((itemId: number): Promise<SPHttpClientResponse> => {
                if (itemId === -1) {
                  throw new Error('No items found in the list');
                }
                latestItemId = itemId;
                this.message(`Loading information about item ID: ${itemId}...`);
             return this._context.spHttpClient.get(`${this._context.pageContext["web"]["absoluteUrl"]}/_api/web/lists/getbytitle('${this._listName}')/items(${latestItemId})?$select=Title,Id`,
                  SPHttpClient.configurations.v1,
                  {  headers: {
                      'Accept': 'application/json;odata=nometadata',
                      'odata-version': ''
                    }
                  });      })
              .then((response: SPHttpClientResponse): Promise<IListItem> => {
                return response.json();
              })
              .then((item: IListItem): void => {
                this.message(`Item ID1: ${item.Id}, Title: ${item.Title}`);
                const body: string = JSON.stringify({
                  'Title': `Updated Item ${new Date()}`
                });
          this._context.spHttpClient.post(`${this._context.pageContext["web"]["absoluteUrl"]}/_api/web/lists/getbytitle('${this._listName}')/items(${item.Id})`,
                  SPHttpClient.configurations.v1,
                  {            headers: {
                      'Accept': 'application/json;odata=nometadata',
                      'Content-type': 'application/json;odata=nometadata',
                      'odata-version': '',
                      'IF-MATCH': '*',
                      'X-HTTP-Method': 'MERGE'
                    },    body: body
                  })
               .then((response: SPHttpClientResponse): void => {
                this.message(`Item with ID: ${latestItemId} successfully updated`);
                  }, (error: any): void => {
                    this.message(`Error updating item: ${error}`);
                  });
              });
          }
        private deleteItem(): void {
            if (!window.confirm('Are you sure you want to delete the latest item?'))      
            {      return;     }
              this.message('Loading latest items...');
              let latestItemId: number = undefined;
              let etag: string = undefined;
              this.getLatestItemId()
               .then((itemId: number): Promise<SPHttpClientResponse> => {
                 if (itemId === -1) {
                  throw new Error('No items found in the list');
               }
              latestItemId = itemId;
             this.message(`Loading information about item ID: ${latestItemId}...`);
              return this._context.spHttpClient.get(`${this._context.pageContext["web"]["absoluteUrl"]}/_api/web/lists/getbytitle('${this._listName}')/items(${latestItemId})?$select=Id`,
               SPHttpClient.configurations.v1,
               {  headers: { 'Accept': 'application/json;odata=nometadata',
                        'odata-version': ''  }
                    });
                })
                .then((response: SPHttpClientResponse): Promise<IListItem> => {
                 etag = response.headers.get('ETag');
                 return response.json();
                })
                .then((item: IListItem): Promise<SPHttpClientResponse> => {
                  this.message(`Deleting item with ID: ${latestItemId}...`);
                  return this._context.spHttpClient.post(`${this._context.pageContext["web"]["absoluteUrl"]}/_api/web/lists/getbytitle('${this._listName}')/items(${item.Id})`,
                 SPHttpClient.configurations.v1,
                    {
                     headers: {
                        'Accept': 'application/json;odata=nometadata',
                        'Content-type': 'application/json;odata=verbose',
                        'odata-version': '',
                       'IF-MATCH': etag,
                        'X-HTTP-Method': 'DELETE'
                      }
                    });      })
                .then((response: SPHttpClientResponse): void => {
                  this.message(`Item with ID: ${latestItemId} successfully deleted`);
                }, (error: any): void => {
                  this.message(`Error deleting item: ${error}`);
                });
            }
                              
}
