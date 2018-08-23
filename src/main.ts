import { WizdomSpfxServices, IWizdomWebApiService, IWizdomCache } from "@wizdom-intranet/services";

export interface IUser {
  accountName: string;
  department: string;
  email: string;
  location: string;
  name: string;
  phone: string;
  pictureUrl: string;
  publicUrl: string;
  extendedProperties: { [key: string]: string };
}

export interface IDepartmentUsersDataService {
  /**
   * Get Users by department (either specific or the user's)
   * @param department The department to get users from
   * @param useUsersDepartment Use the current user's department instead of above
   * @param selectProperties Extra managed properties to get
   * @param resultSource Result source to use
   */
  getUsers: (
    department: string,
    useUsersDepartment: boolean,
    selectProperties: string[],
    resultSource: string
  ) => Promise<IUser[]>;
  /**
   * Get Users by specified search query
   * @param query The search query to use to get users
   * @param selectProperties Extra managed properties to get
   * @param resultSource Result source to use
   */
  getUsersByQuery: (query: string, selectProperties: string[], resultSource: string) => Promise<IUser[]>;
  /**
   * Get Users by specified loginnames
   * @param loginNames The login names of the users (can also be guid of group to retrieve members of)
   * @param selectProperties Extra managed properties to get
   * @param resultSource Result source to use
   */
  getUsersByLoginNames: (loginNames: string[], selectProperties: string[], resultSource: string) => Promise<IUser[]>;
  /**
   * Get list of departments
   */
  getDepartments: () => Promise<string[]>;
}
export default class DepartmentUsersDataService implements IDepartmentUsersDataService {
  constructor(private wizdomServices: WizdomSpfxServices, private webUrl) {}

  public getApiService(): IWizdomWebApiService {
    return this.wizdomServices.WizdomWebApiService;
  }
  public getCache(): IWizdomCache {
    return this.wizdomServices.Cache;
  }
  private baseUrl = "api/wizdom/departmentusers";
  private expireInMilliseconds = 8 * 60 * 60 * 1000; // 8 hours
  private refreshInMilliseconds = 30 * 60 * 1000; // 30 minutes
  private refreshDelayInMilliseconds = 10 * 1000; // 10 seconds

  public getUsers(department: string, useUsersDepartment: boolean, selectProperties: string[], resultSource: string) {
    var cacheKey = `DepartmentUsers.getUsers:${useUsersDepartment ? "[login]" : department}:${selectProperties.join(
      ","
    )}:${resultSource}`;
    return this.getCache().Localstorage.ExecuteCached<IUser[]>(cacheKey, () => {
        var urlPart = [
          `${this.baseUrl}?department=${encodeURIComponent(department)}&useUsersDepartment=${useUsersDepartment}`
        ];
        if (selectProperties) {
          selectProperties.forEach(property => urlPart.push(`&selectProperties=${encodeURIComponent(property)}`));
        }
        if (resultSource) {
          urlPart.push(`&resultSource=${resultSource}`);
        }
        return this.getApiService().Get(urlPart.join("")).then(r => this.fixPictureUrl(<IUser[]>r.users));
      }, this.expireInMilliseconds, this.refreshInMilliseconds, this.refreshDelayInMilliseconds);
    }

  public getUsersByQuery(query: string, selectProperties: string[], resultSource: string) {
    var cacheKey = `DepartmentUsers.getUsersByQuery:${query}:${selectProperties.join(",")}:${resultSource}`;
    return this.getCache().Localstorage.ExecuteCached<IUser[]>(cacheKey, () => {
      var urlPart = [`${this.baseUrl}?query=${encodeURIComponent(query)}`];
        if (selectProperties) {
          selectProperties.forEach(property => urlPart.push(`&selectProperties=${encodeURIComponent(property)}`));
        }
        if (resultSource) {
          urlPart.push(`&resultSource=${resultSource}`);
        }
        return this.getApiService().Get(urlPart.join("")).then(r => this.fixPictureUrl(<IUser[]>r.users));
      }, this.expireInMilliseconds, this.refreshInMilliseconds, this.refreshDelayInMilliseconds);
    }

  public getUsersByLoginNames(loginNames: string[], selectProperties: string[], resultSource: string) {
    var cacheKey = `DepartmentUsers.getUsersByLoginNames:${loginNames.join(",")}:${selectProperties.join(
      ","
    )}:${resultSource}`;
    return this.getCache().Localstorage.ExecuteCached<IUser[]>(cacheKey, () => {
      var urlPart = [`${this.baseUrl}/ensure?`];
        if (selectProperties) {
          selectProperties.forEach(property => urlPart.push(`&selectProperties=${encodeURIComponent(property)}`));
        }
        if (resultSource) {
          urlPart.push(`&resultSource=${resultSource}`);
        }
        return this.getApiService().Post(urlPart.join(""), loginNames).then(r => this.fixPictureUrl(<IUser[]>r.users));
      }, this.expireInMilliseconds, this.refreshInMilliseconds, this.refreshDelayInMilliseconds);
    }

  public getDepartments(): Promise<string[]> {
    var cacheKey = `DepartmentUsers.getDepartments`;
    return this.getCache().Localstorage.ExecuteCached<string[]>(cacheKey, () => {
        return this.getApiService().Get("api/wizdom/365/terms?termId=8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f");
      }, this.expireInMilliseconds, this.refreshInMilliseconds, this.refreshDelayInMilliseconds);
    }

  private fixPictureUrl(users: IUser[]) {
    users.forEach(u => {
      u.pictureUrl = `${this.webUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(
        u.accountName
      )}`;
    });
    return users;
  }
}
