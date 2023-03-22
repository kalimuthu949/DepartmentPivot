import { sp } from "@pnp/sp/presets/all";
import {
  IFilter,
  IListItems,
  IListItemUsingId,
  IAddList,
  IUpdateList,
  ISPList,
  IDetailsListGroup,
  IPeopleObj,
  ISPAttachment,
  IAttachDelete,
  IAttachContents,
  ISPListChoiceField,
} from "./ISPServicesProps";
import { IItemAddResult } from "@pnp/sp/items";
import "@pnp/sp/webs";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";
import { IAttachmentFileInfo } from "@pnp/sp/attachments";

const getAllUsers = async (): Promise<[]> => {
  return await sp.web.siteUsers();
};
const SPAddItem = async (params: IAddList): Promise<IItemAddResult> => {
  return await sp.web.lists
    .getByTitle(params.Listname)
    .items.add(params.RequestJSON);
};
const SPUpdateItem = async (params: IUpdateList): Promise<IItemAddResult> => {
  return await sp.web.lists
    .getByTitle(params.Listname)
    .items.getById(params.ID)
    .update(params.RequestJSON);
};
const SPDeleteItem = async (params: ISPList): Promise<void> => {
  return await sp.web.lists
    .getByTitle(params.Listname)
    .items.getById(params.ID)
    .delete();
};
const SPReadItems = async (params: IListItems): Promise<[]> => {
  params = formatInputs(params);
  let filterValue: string = formatFilterValue(params.Filter);

  return await sp.web.lists
    .getByTitle(params.Listname)
    .items.select(params.Select)
    .filter(filterValue)
    .expand(params.Expand)
    .top(params.Topcount)
    .orderBy(params.Orderby, params.Orderbydecorasc)
    .get();
};

const SPReadItemUsingId = async (params: IListItemUsingId): Promise<[]> => {
  return await sp.web.lists
    .getByTitle(params.Listname)
    .items.getById(params.SelectedId)
    .select(params.Select)
    .expand(params.Expand)
    .get();
};

const SPAddAttachments = async (params: ISPAttachment) => {
  const files: any[] = params.Attachments;
  return await sp.web.lists
    .getByTitle(params.ListName)
    .items.getById(params.ListID)
    .attachmentFiles.addMultiple(files);
};

const SPGetAttachments = async (params: ISPList) => {
  const item: any = sp.web.lists
    .getByTitle(params.Listname)
    .items.getById(params.ID);
  return await item.attachmentFiles();
};

const SPDeleteAttachments = async (params: IAttachDelete) => {
  return await sp.web.lists
    .getByTitle(params.ListName)
    .items.getById(params.ListID)
    .attachmentFiles.getByName(params.AttachmentName)
    .delete();
};

const SPGetChoices = async (params: ISPListChoiceField) => {
  return await sp.web.lists
    .getByTitle(params.Listname)
    .fields.getByInternalNameOrTitle(params.FieldName)
    .get();
};

const SPDetailsListGroupItems = async (params: IDetailsListGroup) => {
  let newRecords = [];
  params.Data.forEach((arr, index) => {
    newRecords.push({
      Lesson: arr[params.Column],
      indexValue: index,
    });
  });

  let varGroup = [];
  let UniqueRecords = newRecords.reduce(function (item, e1) {
    var matches = item.filter(function (e2) {
      return e1[params.Column] === e2[params.Column];
    });

    if (matches.length == 0) {
      item.push(e1);
    }
    return item;
  }, []);

  UniqueRecords.forEach((ur) => {
    let recordLength = newRecords.filter((arr) => {
      return arr[params.Column] == ur[params.Column];
    }).length;
    varGroup.push({
      key: ur[params.Column],
      name: ur[params.Column],
      startIndex: ur.indexValue,
      count: recordLength,
    });
  });
  // console.log([...varGroup]);
  return varGroup;
};

const readItemsFromSharepointListForDashbaord = async (
  params: IListItems
): Promise<[]> => {
  params = formatInputs(params);
  let filterValue: string = formatFilterValue(params.Filter);
  let skipcount = params.PageNumber * params.PageCount - params.PageCount;

  return await sp.web.lists
    .getByTitle(params.Listname)
    .items.select(params.Select)
    .filter(filterValue)
    .expand(params.Expand)
    .skip(skipcount)
    .top(params.PageCount)
    .orderBy(params.Orderby, params.Orderbydecorasc)
    .get();
};

const formatInputs = (data: IListItems): IListItems => {
  !data.Select ? (data.Select = "*") : "";
  !data.Topcount ? (data.Topcount = 5000) : "";
  !data.Orderby ? (data.Orderby = "ID") : "";
  !data.Expand ? (data.Expand = "") : "";
  !data.Orderbydecorasc == true && !data.Orderbydecorasc == false
    ? (data.Orderbydecorasc = true)
    : "";
  !data.PageCount ? (data.PageCount = 10) : "";
  !data.PageNumber ? (data.PageNumber = 1) : "";

  return data;
};
const formatFilterValue = (params: IFilter[]): string => {
  let strFilter: string = "";
  if (params) {
    for (let i = 0; i < params.length; i++) {
      if (params[i].FilterKey) {
        if (i != 0) {
          strFilter += " and ";
        }

        if (
          params[i].Operator.toLocaleLowerCase() == "eq" ||
          params[i].Operator.toLocaleLowerCase() == "neq" ||
          params[i].Operator.toLocaleLowerCase() == "gt" ||
          params[i].Operator.toLocaleLowerCase() == "lt" ||
          params[i].Operator.toLocaleLowerCase() == "ge" ||
          params[i].Operator.toLocaleLowerCase() == "le"
        )
          strFilter +=
            params[i].FilterKey +
            " " +
            params[i].Operator +
            "'" +
            params[i].FilterValue +
            "'";
        else if (params[i].Operator.toLocaleLowerCase() == "substringof")
          strFilter +=
            params[i].Operator +
            "('" +
            params[i].FilterKey +
            "','" +
            params[i].FilterValue +
            "')";
      }
    }
  }
  return strFilter;
};

export default {
  getAllUsers,
  SPAddItem,
  SPUpdateItem,
  SPDeleteItem,
  SPReadItems,
  SPDetailsListGroupItems,
  SPGetChoices,
  SPAddAttachments,
  SPGetAttachments,
  SPDeleteAttachments,
  SPReadItemUsingId,
};