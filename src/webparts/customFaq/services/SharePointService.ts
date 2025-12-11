import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

/**
 * Interface for SharePoint list information
 */
export interface IListInfo {
  id: string;
  title: string;
}

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
  public async getLists(): Promise<IListInfo[]> {
    const endpoint = `${this._siteUrl}/_api/web/lists?$filter=Hidden eq false and BaseTemplate eq 100&$select=Id,Title&$orderby=Title`;

    try {
      const response: SPHttpClientResponse = await this._context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Error fetching lists: ${response.statusText}`);
      }

      const data = await response.json();

      return data.value.map((list: { Id: string; Title: string }) => ({
        id: list.Id,
        title: list.Title
      }));
    } catch (error) {
      console.error('SharePointService.getLists error:', error);
      throw error;
    }
  }

  /**
   * Get columns for a specific list
   * Returns only Text and Note (multi-line text) columns
   */
  public async getListColumns(listId: string): Promise<IColumnInfo[]> {
    // Filter for Text and Note field types
    const endpoint = `${this._siteUrl}/_api/web/lists(guid'${listId}')/fields?$filter=(TypeAsString eq 'Text' or TypeAsString eq 'Note') and Hidden eq false and ReadOnlyField eq false&$select=Id,InternalName,Title,TypeAsString&$orderby=Title`;

    try {
      const response: SPHttpClientResponse = await this._context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Error fetching columns: ${response.statusText}`);
      }

      const data = await response.json();

      return data.value.map((field: { Id: string; InternalName: string; Title: string; TypeAsString: string }) => ({
        id: field.Id,
        internalName: field.InternalName,
        title: field.Title,
        type: field.TypeAsString
      }));
    } catch (error) {
      console.error('SharePointService.getListColumns error:', error);
      throw error;
    }
  }

  /**
   * Get all items from a list with specified columns
   * Also fetches attachments for each item
   */
  public async getListItems(
    listId: string,
    titleColumn: string,
    descriptionColumn: string
  ): Promise<IFaqItem[]> {
    // Build select query - always include Id and Attachments
    const selectFields = ['Id', titleColumn];
    if (descriptionColumn !== titleColumn) {
      selectFields.push(descriptionColumn);
    }
    selectFields.push('Attachments');

    const endpoint = `${this._siteUrl}/_api/web/lists(guid'${listId}')/items?$select=${selectFields.join(',')}&$top=100`;

    try {
      const response: SPHttpClientResponse = await this._context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Error fetching items: ${response.statusText}`);
      }

      const data = await response.json();

      // Process items and fetch attachments
      const items: IFaqItem[] = await Promise.all(
        data.value.map(async (item: Record<string, unknown>) => {
          const faqItem: IFaqItem = {
            id: item.Id as number,
            title: (item[titleColumn] as string) || '',
            description: (item[descriptionColumn] as string) || '',
            attachments: []
          };

          // If item has attachments, fetch them
          if (item.Attachments) {
            faqItem.attachments = await this._getItemAttachments(listId, item.Id as number);
          }

          return faqItem;
        })
      );

      // Filter out items without titles
      return items.filter(item => item.title.trim() !== '');
    } catch (error) {
      console.error('SharePointService.getListItems error:', error);
      throw error;
    }
  }

  /**
   * Get attachments for a specific list item
   */
  private async _getItemAttachments(listId: string, itemId: number): Promise<IAttachment[]> {
    const endpoint = `${this._siteUrl}/_api/web/lists(guid'${listId}')/items(${itemId})/AttachmentFiles`;

    try {
      const response: SPHttpClientResponse = await this._context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.warn(`Could not fetch attachments for item ${itemId}`);
        return [];
      }

      const data = await response.json();

      return data.value.map((attachment: { FileName: string; ServerRelativeUrl: string }) => ({
        fileName: attachment.FileName,
        url: `${this._context.pageContext.web.absoluteUrl.split('/').slice(0, 3).join('/')}${attachment.ServerRelativeUrl}`
      }));
    } catch (error) {
      console.warn('Error fetching attachments:', error);
      return [];
    }
  }

  /**
   * Get a specific list by ID
   */
  public async getListById(listId: string): Promise<IListInfo | null> {
    const endpoint = `${this._siteUrl}/_api/web/lists(guid'${listId}')?$select=Id,Title`;

    try {
      const response: SPHttpClientResponse = await this._context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        return null;
      }

      const data = await response.json();

      return {
        id: data.Id,
        title: data.Title
      };
    } catch (error) {
      console.warn('Error fetching list:', error);
      return null;
    }
  }

  /**
   * Check if a column exists in a list
   */
  public async columnExists(listId: string, columnInternalName: string): Promise<boolean> {
    const endpoint = `${this._siteUrl}/_api/web/lists(guid'${listId}')/fields?$filter=InternalName eq '${columnInternalName}'&$select=Id`;

    try {
      const response: SPHttpClientResponse = await this._context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        return false;
      }

      const data = await response.json();
      return data.value && data.value.length > 0;
    } catch (error) {
      console.warn('Error checking column existence:', error);
      return false;
    }
  }
}
