import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

/**
 * Interface for SharePoint list information
 */
export interface IListInfo {
  id: string;
  title: string;
}

//Vishnu to fix security issue based on scann
/**
 * Interface for SharePoint column information
 */
export interface IColumnInfo {
  id: string;
  internalName: string;
  title: string;
  type: string;
}

/**
 * Interface for attachment information
 */
export interface IAttachment {
  fileName: string;
  url: string;
}

/**
 * Interface for FAQ item
 */
export interface IFaqItem {
  id: number;
  title: string;
  description: string;
  category: string;
  attachments: IAttachment[];
}

/**
 * SharePoint Service - Helper methods for retrieving lists, columns, and items
 */
export class SharePointService {
  private _context: WebPartContext;
  private _siteUrl: string;

  constructor(context: WebPartContext) {
    this._context = context;
    this._siteUrl = context.pageContext.web.absoluteUrl;
  }

  /**
   * Get all lists from the current site
   * Filters out hidden lists and system lists
   */
  public getLists(): Promise<IListInfo[]> {
    const endpoint = this._siteUrl + '/_api/web/lists?$filter=Hidden eq false and BaseTemplate eq 100&$select=Id,Title&$orderby=Title';

    return this._context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    )
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error('Error fetching lists: ' + response.statusText);
        }
        return response.json();
      })
      .then((data: { value: Array<{ Id: string; Title: string }> }) => {
        const lists: IListInfo[] = [];
        for (let i = 0; i < data.value.length; i++) {
          lists.push({
            id: data.value[i].Id,
            title: data.value[i].Title
          });
        }
        return lists;
      });
  }

  /**
   * Get columns for a specific list
   * Returns Text, Note, and Choice columns
   */
  public getListColumns(listId: string): Promise<IColumnInfo[]> {
    // Filter for Text, Note, and Choice field types
    const endpoint = this._siteUrl + '/_api/web/lists(guid\'' + listId + '\')/fields?$filter=(TypeAsString eq \'Text\' or TypeAsString eq \'Note\' or TypeAsString eq \'Choice\') and Hidden eq false and ReadOnlyField eq false&$select=Id,InternalName,Title,TypeAsString&$orderby=Title';

    return this._context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    )
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error('Error fetching columns: ' + response.statusText);
        }
        return response.json();
      })
      .then((data: { value: Array<{ Id: string; InternalName: string; Title: string; TypeAsString: string }> }) => {
        const columns: IColumnInfo[] = [];
        for (let i = 0; i < data.value.length; i++) {
          columns.push({
            id: data.value[i].Id,
            internalName: data.value[i].InternalName,
            title: data.value[i].Title,
            type: data.value[i].TypeAsString
          });
        }
        return columns;
      });
  }

  /**
   * Get all items from a list with specified columns
   * Also fetches attachments for each item
   */
  public getListItems(
    listId: string,
    titleColumn: string,
    descriptionColumn: string,
    categoryColumn?: string
  ): Promise<IFaqItem[]> {
    // Build select query - always include Id and Attachments
    const selectFields: string[] = ['Id', titleColumn];
    if (descriptionColumn !== titleColumn) {
      selectFields.push(descriptionColumn);
    }
    if (categoryColumn && categoryColumn !== titleColumn && categoryColumn !== descriptionColumn) {
      selectFields.push(categoryColumn);
    }
    selectFields.push('Attachments');

    const endpoint = this._siteUrl + '/_api/web/lists(guid\'' + listId + '\')/items?$select=' + selectFields.join(',') + '&$top=500';

    const self = this;

    return this._context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    )
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error('Error fetching items: ' + response.statusText);
        }
        return response.json();
      })
      .then((data: { value: Array<{ [key: string]: unknown }> }) => {
        // Process items and fetch attachments
        const itemPromises: Array<Promise<IFaqItem>> = [];

        for (let i = 0; i < data.value.length; i++) {
          const item = data.value[i];
          itemPromises.push(self._processItem(item, listId, titleColumn, descriptionColumn, categoryColumn));
        }

        return Promise.all(itemPromises);
      })
      .then((items: IFaqItem[]) => {
        // Filter out items without titles
        const filteredItems: IFaqItem[] = [];
        for (let i = 0; i < items.length; i++) {
          if (items[i].title && items[i].title.trim() !== '') {
            filteredItems.push(items[i]);
          }
        }
        return filteredItems;
      });
  }

  /**
   * Process a single item and fetch its attachments
   */
  private _processItem(
    item: { [key: string]: unknown },
    listId: string,
    titleColumn: string,
    descriptionColumn: string,
    categoryColumn?: string
  ): Promise<IFaqItem> {
    const self = this;
    const faqItem: IFaqItem = {
      id: item.Id as number,
      title: (item[titleColumn] as string) || '',
      description: (item[descriptionColumn] as string) || '',
      category: categoryColumn ? ((item[categoryColumn] as string) || '') : '',
      attachments: []
    };

    // If item has attachments, fetch them
    if (item.Attachments) {
      return self._getItemAttachments(listId, item.Id as number)
        .then((attachments: IAttachment[]) => {
          faqItem.attachments = attachments;
          return faqItem;
        });
    }

    return Promise.resolve(faqItem);
  }

  /**
   * Get attachments for a specific list item
   */
  private _getItemAttachments(listId: string, itemId: number): Promise<IAttachment[]> {
    const endpoint = this._siteUrl + '/_api/web/lists(guid\'' + listId + '\')/items(' + itemId + ')/AttachmentFiles';
    const self = this;

    return this._context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    )
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          console.warn('Could not fetch attachments for item ' + itemId);
          return { value: [] };
        }
        return response.json();
      })
      .then((data: { value: Array<{ FileName: string; ServerRelativeUrl: string }> }) => {
        const attachments: IAttachment[] = [];
        const baseUrl = self._context.pageContext.web.absoluteUrl.split('/').slice(0, 3).join('/');

        for (let i = 0; i < data.value.length; i++) {
          attachments.push({
            fileName: data.value[i].FileName,
            url: baseUrl + data.value[i].ServerRelativeUrl
          });
        }

        return attachments;
      })
      .catch((error: Error) => {
        console.warn('Error fetching attachments:', error);
        return [];
      });
  }

  /**
   * Get a specific list by ID
   */
  public getListById(listId: string): Promise<IListInfo | null> {
    const endpoint = this._siteUrl + '/_api/web/lists(guid\'' + listId + '\')?$select=Id,Title';

    return this._context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    )
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return null;
        }
        return response.json();
      })
      .then((data: { Id: string; Title: string } | null) => {
        if (!data) {
          return null;
        }
        return {
          id: data.Id,
          title: data.Title
        };
      })
      .catch(() => {
        return null;
      });
  }

  /**
   * Check if a column exists in a list
   */
  public columnExists(listId: string, columnInternalName: string): Promise<boolean> {
    const endpoint = this._siteUrl + '/_api/web/lists(guid\'' + listId + '\')/fields?$filter=InternalName eq \'' + columnInternalName + '\'&$select=Id';

    return this._context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    )
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return false;
        }
        return response.json();
      })
      .then((data: { value: Array<{ Id: string }> }) => {
        return data.value && data.value.length > 0;
      })
      .catch(() => {
        return false;
      });
  }
}