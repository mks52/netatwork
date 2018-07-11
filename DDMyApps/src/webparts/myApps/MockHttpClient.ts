import {MyAppsList} from './MyAppsWebPart';


export default class MockHttpClient {
    private static _items: MyAppsList[] = [
        {Title: "Title 1", TypeofApp: "O365 App", O365App: "Delve", ThirdPartyAppImageLink: {Url: "Url1"}, ThirdPartyAppUrl: {Url: "Link 1"}},
        {Title: "Title 1", TypeofApp: "O365 App", O365App: "Delve", ThirdPartyAppImageLink: {Url: "Url1"}, ThirdPartyAppUrl: {Url: "Link 1"}}
    ];
    public static get(restUrl: string, options?: any): Promise<MyAppsList[]> {
        return new Promise<MyAppsList[]>((resolve) => {
            resolve(MockHttpClient._items);
        });
    }
}