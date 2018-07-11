import {VideosList} from './VideosWebPart';


export default class MockHttpClient {
    private static _items: VideosList[] = [
        {ChannelId: "ChannelId 1", Description: "Description 1", DisplayFormUrl: "Link 1", FileName: "Title 1", ThumbnailUrl: "image 1"},
        {ChannelId: "ChannelId 2", Description: "Description 2", DisplayFormUrl: "Link 2", FileName: "Title 2", ThumbnailUrl: "image 2"}
    ];
    public static get(restUrl: string, options?: any): Promise<VideosList[]> {
        return new Promise<VideosList[]>((resolve) => {
            resolve(MockHttpClient._items);
        });
    }
}