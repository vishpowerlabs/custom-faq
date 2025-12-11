import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  type IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webparts-base';
import {
  ThemeProvider,
  ThemeChangedEventArgs,
  IReadonlyTheme
} from '@microsoft/sp-component-base';

import styles from './CustomFaq.module.scss';
import { SharePointService, IListInfo, IColumnInfo, IFaqItem } from './services/SharePointService';

export interface ICustomFaqWebPartProps {
  webpartTitle: string;
  webpartDescription: string;
  selectedList: string;
  titleColumn: string;
  descriptionColumn: string;
  allowMultipleExpanded: boolean;
}

export default class CustomFaqWebPart extends BaseClientSideWebPart<ICustomFaqWebPartProps> {

  private _spService: SharePointService;
  private _lists: IListInfo[] = [];
  private _columns: IColumnInfo[] = [];
  private _faqItems: IFaqItem[] = [];
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;

  /**
   * Initialize the web part
   */
  protected async onInit(): Promise<void> {
    await super.onInit();

    // Initialize SharePoint service
    this._spService = new SharePointService(this.context);

    // Consume the ThemeProvider service for section background support
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

    // Get the current theme variant
    this._themeVariant = this._themeProvider.tryGetTheme();

    // Apply CSS variables from theme
    if (this._themeVariant) {
      this._setCSSVariables(this._themeVariant);
    }

    // Register handler for theme changes
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);

    // Load initial data
    await this._loadLists();

    if (this.properties.selectedList) {
      await this._loadColumns();
      await this._loadFaqItems();
    }
  }

  /**
   * Handle theme change events (section background changes)
   */
  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    if (this._themeVariant) {
      this._setCSSVariables(this._themeVariant);
    }
    this.render();
  }

  /**
   * Convert theme semantic colors to CSS variables for use in SCSS
   */
  private _setCSSVariables(theme: IReadonlyTheme): void {
    if (!this.domElement) return;

    // Set semantic colors as CSS variables
    if (theme.semanticColors) {
      const semanticColors = theme.semanticColors as Record<string, string>;
      for (const [key, value] of Object.entries(semanticColors)) {
        if (value) {
          this.domElement.style.setProperty(`--${key}`, value);
        }
      }
    }

    // Set palette colors as CSS variables
    if (theme.palette) {
      const palette = theme.palette as Record<string, string>;
      for (const [key, value] of Object.entries(palette)) {
        if (value) {
          this.domElement.style.setProperty(`--${key}`, value);
        }
      }
    }
  }

  /**
   * Render the web part
   */
  public render(): void {
    // Build inline styles from theme for elements that need direct styling
    const semanticColors = this._themeVariant?.semanticColors;
    const palette = this._themeVariant?.palette;

    // Header background using theme primary color
    const headerBgColor = palette?.themePrimary || '#0078d4';
    const headerBgColorDark = palette?.themeDark || '#005a9e';

    // Body colors
    const bodyBackground = semanticColors?.bodyBackground || '#ffffff';
    const bodyText = semanticColors?.bodyText || '#323130';
    const bodySubtext = semanticColors?.bodySubtext || '#605e5c';

    // Interactive colors
    const linkColor = semanticColors?.link || palette?.themePrimary || '#0078d4';
    const linkHoverColor = semanticColors?.linkHovered || palette?.themeDark || '#005a9e';

    // Border and divider colors
    const bodyDivider = semanticColors?.bodyDivider || '#edebe9';

    // Hover states
    const listItemBackgroundHovered = semanticColors?.listItemBackgroundHovered || '#f3f2f1';

    // Card/container background
    const cardBackground = semanticColors?.cardStandoutBackground || bodyBackground;

    // Attachment area background
    const neutralLighter = palette?.neutralLighter || '#f4f4f4';
    const neutralTertiary = palette?.neutralTertiary || '#a6a6a6';

    let faqItemsHtml = '';

    if (!this.properties.selectedList) {
      faqItemsHtml = `
        <div class="${styles.emptyState}" style="color: ${bodySubtext};">
          <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="${neutralTertiary}" stroke-width="1.5">
            <circle cx="12" cy="12" r="10"/>
            <path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"/>
            <line x1="12" y1="17" x2="12.01" y2="17"/>
          </svg>
          <p>Please configure the web part by selecting a list from the property pane.</p>
        </div>
      `;
    } else if (this._faqItems.length === 0) {
      faqItemsHtml = `
        <div class="${styles.emptyState}" style="color: ${bodySubtext};">
          <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="${neutralTertiary}" stroke-width="1.5">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
            <polyline points="14 2 14 8 20 8"/>
            <line x1="12" y1="18" x2="12" y2="12"/>
            <line x1="9" y1="15" x2="15" y2="15"/>
          </svg>
          <p>No FAQ items found in the selected list.</p>
        </div>
      `;
    } else {
      faqItemsHtml = this._faqItems.map((item, index) => {
        let attachmentsHtml = '';

        if (item.attachments && item.attachments.length > 0) {
          const attachmentLinks = item.attachments.map(att => `
            <a href="${att.url}" 
               target="_blank" 
               rel="noopener noreferrer" 
               class="${styles.attachmentLink}"
               style="color: ${linkColor};">
              <span class="${styles.attachmentIcon}">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                  <polyline points="14 2 14 8 20 8"/>
                </svg>
              </span>
              ${this._escapeHtml(att.fileName)}
            </a>
          `).join('');

          attachmentsHtml = `
            <div class="${styles.attachments}" style="background-color: ${neutralLighter};">
              <div class="${styles.attachmentsLabel}" style="color: ${neutralTertiary};">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                  <path d="M21.44 11.05l-9.19 9.19a6 6 0 0 1-8.49-8.49l9.19-9.19a4 4 0 0 1 5.66 5.66l-9.2 9.19a2 2 0 0 1-2.83-2.83l8.49-8.48"/>
                </svg>
                Attachments
              </div>
              ${attachmentLinks}
            </div>
          `;
        }

        return `
          <div class="${styles.faqItem}" data-index="${index}">
            <div class="${styles.faqQuestion}" 
                 style="border-bottom-color: ${bodyDivider};"
                 data-hover-bg="${listItemBackgroundHovered}">
              <span class="${styles.faqQuestionText}" style="color: ${bodyText};">
                ${this._escapeHtml(item.title)}
              </span>
              <span class="${styles.faqChevron}" style="color: ${bodySubtext};">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                  <path d="M6 9l6 6 6-6"/>
                </svg>
              </span>
            </div>
            <div class="${styles.faqAnswer}">
              <div class="${styles.faqAnswerContent}" style="color: ${bodySubtext};">
                ${this._formatDescription(item.description)}
              </div>
              ${attachmentsHtml}
            </div>
          </div>
        `;
      }).join('');
    }

    // Build header section
    const headerHtml = (this.properties.webpartTitle || this.properties.webpartDescription) ? `
      <div class="${styles.faqHeader}" 
           style="background: linear-gradient(135deg, ${headerBgColor} 0%, ${headerBgColorDark} 100%);">
        ${this.properties.webpartTitle ? `<h2 class="${styles.faqTitle}">${this._escapeHtml(this.properties.webpartTitle)}</h2>` : ''}
        ${this.properties.webpartDescription ? `<p class="${styles.faqDescription}">${this._escapeHtml(this.properties.webpartDescription)}</p>` : ''}
      </div>
    ` : '';

    this.domElement.innerHTML = `
      <div class="${styles.customFaq}" style="background-color: ${cardBackground};">
        ${headerHtml}
        <div class="${styles.faqList}">
          ${faqItemsHtml}
        </div>
      </div>
    `;

    // Attach event listeners for accordion functionality
    this._attachEventListeners();
  }

  /**
   * Attach click event listeners to FAQ items
   */
  private _attachEventListeners(): void {
    const faqItems = this.domElement.querySelectorAll(`.${styles.faqItem}`);
    const questions = this.domElement.querySelectorAll(`.${styles.faqQuestion}`);

    questions.forEach((question, index) => {
      // Add hover effect
      question.addEventListener('mouseenter', () => {
        const hoverBg = question.getAttribute('data-hover-bg');
        if (hoverBg) {
          (question as HTMLElement).style.backgroundColor = hoverBg;
        }
      });

      question.addEventListener('mouseleave', () => {
        (question as HTMLElement).style.backgroundColor = '';
      });

      // Add click handler for accordion
      question.addEventListener('click', () => {
        const faqItem = faqItems[index];

        if (!this.properties.allowMultipleExpanded) {
          // Close all other items
          faqItems.forEach((item, i) => {
            if (i !== index && item.classList.contains(styles.expanded)) {
              item.classList.remove(styles.expanded);
            }
          });
        }

        // Toggle current item
        faqItem.classList.toggle(styles.expanded);
      });
    });
  }

  /**
   * Escape HTML to prevent XSS
   */
  private _escapeHtml(text: string): string {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  /**
   * Format description text (handle rich text or plain text)
   */
  private _formatDescription(description: string): string {
    if (!description) return '';

    // Check if it's HTML content (from multi-line rich text field)
    if (description.includes('<') && description.includes('>')) {
      // Sanitize but allow basic formatting tags
      return description;
    }

    // Plain text - convert line breaks to <br>
    return this._escapeHtml(description).replace(/\n/g, '<br>');
  }

  /**
   * Load all lists from the current site
   */
  private async _loadLists(): Promise<void> {
    try {
      this._lists = await this._spService.getLists();
    } catch (error) {
      console.error('Error loading lists:', error);
      this._lists = [];
    }
  }

  /**
   * Load columns for the selected list
   */
  private async _loadColumns(): Promise<void> {
    if (!this.properties.selectedList) {
      this._columns = [];
      return;
    }

    try {
      this._columns = await this._spService.getListColumns(this.properties.selectedList);
    } catch (error) {
      console.error('Error loading columns:', error);
      this._columns = [];
    }
  }

  /**
   * Load FAQ items from the selected list
   */
  private async _loadFaqItems(): Promise<void> {
    if (!this.properties.selectedList || !this.properties.titleColumn || !this.properties.descriptionColumn) {
      this._faqItems = [];
      return;
    }

    try {
      this._faqItems = await this._spService.getListItems(
        this.properties.selectedList,
        this.properties.titleColumn,
        this.properties.descriptionColumn
      );
    } catch (error) {
      console.error('Error loading FAQ items:', error);
      this._faqItems = [];
    }
  }

  /**
   * Get data version
   */
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /**
   * Handle property pane field changes
   */
  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): Promise<void> {
    if (propertyPath === 'selectedList' && newValue !== oldValue) {
      // Reset column selections when list changes
      this.properties.titleColumn = '';
      this.properties.descriptionColumn = '';
      this._faqItems = [];

      // Load columns for new list
      await this._loadColumns();

      // Refresh property pane to show new columns
      this.context.propertyPane.refresh();
    }

    if ((propertyPath === 'titleColumn' || propertyPath === 'descriptionColumn') && newValue !== oldValue) {
      // Reload FAQ items when columns change
      await this._loadFaqItems();
    }

    this.render();
  }

  /**
   * Get dropdown options for lists
   */
  private _getListOptions(): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = [
      { key: '', text: '-- Select a list --' }
    ];

    this._lists.forEach(list => {
      options.push({ key: list.id, text: list.title });
    });

    return options;
  }

  /**
   * Get dropdown options for title columns (text columns only)
   */
  private _getTitleColumnOptions(): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = [
      { key: '', text: '-- Select a column --' }
    ];

    this._columns
      .filter(col => col.type === 'Text' || col.type === 'Note')
      .forEach(col => {
        options.push({ key: col.internalName, text: col.title });
      });

    return options;
  }

  /**
   * Get dropdown options for description columns (text and multi-line text columns)
   */
  private _getDescriptionColumnOptions(): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = [
      { key: '', text: '-- Select a column --' }
    ];

    this._columns
      .filter(col => col.type === 'Text' || col.type === 'Note')
      .forEach(col => {
        options.push({ key: col.internalName, text: `${col.title} (${col.type === 'Note' ? 'Multi-line' : 'Single-line'})` });
      });

    return options;
  }

  /**
   * Configure the property pane
   */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configure the FAQ web part settings'
          },
          groups: [
            {
              groupName: 'Display Settings',
              groupFields: [
                PropertyPaneTextField('webpartTitle', {
                  label: 'Webpart Title',
                  placeholder: 'Enter a title for the FAQ section',
                  value: this.properties.webpartTitle
                }),
                PropertyPaneTextField('webpartDescription', {
                  label: 'Webpart Description',
                  placeholder: 'Enter a description',
                  multiline: true,
                  rows: 3,
                  value: this.properties.webpartDescription
                })
              ]
            },
            {
              groupName: 'Data Source',
              groupFields: [
                PropertyPaneDropdown('selectedList', {
                  label: 'Select List',
                  options: this._getListOptions(),
                  selectedKey: this.properties.selectedList
                }),
                PropertyPaneDropdown('titleColumn', {
                  label: 'Title Column',
                  options: this._getTitleColumnOptions(),
                  selectedKey: this.properties.titleColumn,
                  disabled: !this.properties.selectedList
                }),
                PropertyPaneDropdown('descriptionColumn', {
                  label: 'Description Column',
                  options: this._getDescriptionColumnOptions(),
                  selectedKey: this.properties.descriptionColumn,
                  disabled: !this.properties.selectedList
                })
              ]
            },
            {
              groupName: 'Accordion Behavior',
              groupFields: [
                PropertyPaneToggle('allowMultipleExpanded', {
                  label: 'Allow Multiple Items Expanded',
                  onText: 'Yes',
                  offText: 'No',
                  checked: this.properties.allowMultipleExpanded
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
