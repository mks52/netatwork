import {NewsList} from './NewsWebPart';


export default class MockHttpClient {
    private static _items: NewsList[] = [
        {Title: "Title 1", Image: {Url: "Thumbnail 1"}, Description: "Description 1", Link: {Url: "thumbnail 1"}, Display: true},
        {Title: "Title 1", Image: {Url: "Thumbnail 2"}, Description: "Description 2", Link: {Url: "thumbnail 2"}, Display: true}
    ];
    public static get(restUrl: string, options?: any): Promise<NewsList[]> {
        return new Promise<NewsList[]>((resolve) => {
            resolve(MockHttpClient._items);
        });
    }
}