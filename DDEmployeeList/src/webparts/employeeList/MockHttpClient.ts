import {EmployeeList} from './EmployeeListWebPart';


export default class MockHttpClient {
    private static _items: EmployeeList[] = [
        {FirstName: "John", LastName: "Doe", Picture: {Url: "picture url"},  BirthMonth: 5, BirthDay: 2, JoiningDay: 5, JoiningMonth: 5, LastWorkingDay: new Date("1968-11-16"), Department: "Department", WorkCity: "New York", WorkState: "NY"},
        {FirstName: "Jane", LastName: "Doe", Picture: {Url: "picture url"},  BirthMonth: 5, BirthDay: 2, JoiningDay: 5, JoiningMonth: 5, LastWorkingDay: new Date("1968-11-16"), Department: "Department", WorkCity: "New York", WorkState: "NY"}];
    public static get(restUrl: string, options?: any): Promise<EmployeeList[]> {
        return new Promise<EmployeeList[]>((resolve) => {
            resolve(MockHttpClient._items);
        });
    }
}