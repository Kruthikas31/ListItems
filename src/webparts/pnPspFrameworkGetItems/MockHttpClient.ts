import { ISPList } from './PnPspFrameworkGetItemsWebPart';
 
export default class MockHttpClient {
   private static _items: ISPList[] = [{ Title: '123', EmployeeName: 'John', Experience: 0,Branch: 'Bangalore' },];
   public static get(restUrl: string, options?: any): Promise<ISPList[]> {
     return new Promise<ISPList[]>((resolve) => {
           resolve(MockHttpClient._items);
       });
   }
 }