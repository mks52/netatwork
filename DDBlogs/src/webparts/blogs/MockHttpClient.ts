import {BlogsList} from './BlogsWebPart';


export default class MockHttpClient {
    private static _items: BlogsList[] = [
        {Title: "Title 1", BlogUrl: {Description: "Url 1"}, PersonId: 1, ImageUrl: {Description: "thumbnail 1"}, Excerpts: "Description 1", Display: true},
        {Title: "Title 2", BlogUrl: {Description: "Url 2"}, PersonId: 2, ImageUrl: {Description: "thumbnail 1"}, Excerpts: "Description 2", Display: true}
    ];
    public static get(restUrl: string, options?: any): Promise<BlogsList[]> {
        return new Promise<BlogsList[]>((resolve) => {
            resolve(MockHttpClient._items);
        });
    }
}