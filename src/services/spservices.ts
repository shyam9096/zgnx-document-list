import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Web } from "@pnp/sp/webs";
let Siteurl: any;
// Class Services
export default class spservices {
  constructor(private context: any, SiteUrl) {
    // Setuo Context to PnPjs and MSGraph
    sp.setup({
      spfxContext: this.context,
    });
  }

  public async getListItem(
    Siteurl,
    listId: any,
    viewFields: any,
    expandFields: any
  ) {
    if (Siteurl != "" && Siteurl != undefined) {
      try {
        var event = await Web(Siteurl)
          .getFolderByServerRelativeUrl(listId)
          .select(viewFields)
          .expand("Folders, Files,Editor/Id")
          .get();

        return event;
      } catch (err) {
        return [];
      }
    } else {
      try {
        var event = await sp.web
          .getFolderByServerRelativeUrl(listId)
          .select(viewFields)
          .expand("Folders, Files,LinkURL")
          .get();

        return event;
      } catch (err) {
        return [];
      }
    }
  }

  public async getListDetails(listId: any, Siteurl) {
    if (Siteurl != "" && Siteurl != undefined) {
      try {
        var event = await Web(Siteurl).lists.getById(listId).get();
        return event;
      } catch (err) {
        return [];
      }
    } else {
      try {
        var event = await sp.web.lists.getById(listId).get();
        return event;
      } catch (err) {
        return [];
      }
    }
  }

  public async addListItem(listName: any, eventData: any) {
    return await sp.web.lists.getByTitle(listName).items.add(eventData);
  }

  public async getSiteLists(Siteurl) {
    let results: any[] = [];
    console.log(Siteurl);
    if (Siteurl != "" && Siteurl != undefined) {
      try {
        results = await Web(Siteurl)
          .lists.select("Title", "ID")
          .filter("BaseTemplate eq 101")
          .get();
      } catch (error) {
        return Promise.reject(error);
      }
      return results;
    } else {
      try {
        results = await sp.web.lists
          .select("Title", "ID")
          .filter("BaseTemplate eq 101")
          .get();
      } catch (error) {
        return Promise.reject(error);
      }
      return results;
    }
  }
}
