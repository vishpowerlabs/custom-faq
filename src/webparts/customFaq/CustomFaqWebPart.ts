import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';

import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
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
  categoryColumn: string;
  allowMultipleExpanded: boolean;
  webpartTitleFontSize: string;
  webpartDescriptionFontSize: string;
  faqTitleFontSize: string;
  faqDescriptionFontSize: string;
}

export default class CustomFaqWebPart extends BaseClientSideWebPart<ICustomFaqWebPartProps> {

  private _spService: SharePointService;
  private _lists: IListInfo[] = [];
  private _columns: IColumnInfo[] = [];
  private _faqItems: IFaqItem[] = [];
  private _themeProvider: ThemeProvider | undefined;
  private _themeVariant: IReadonlyTheme | undefined;
  private _selectedCategory: string = 'All';
  private _categories: string[] = [];

  /**
   * Initialize the web part
   */
  protected onInit(): Promise<void> {
    return super.onInit().then((): void => {
      // Initialize SharePoint service
      this._spService = new SharePointService(this.context as WebPartContext);

      // Try to consume the ThemeProvider service for section background support
      try {
        this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
        if (this._themeProvider) {
          // Get the current theme variant
          this._themeVariant = this._themeProvider.tryGetTheme();

          // Register handler for theme changes
          this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);
        }
      } catch (error) {
        console.log('ThemeProvider not available:', error);
      }

      // Apply CSS variables from theme
      if (this._themeVariant) {
        this._setCSSVariables(this._themeVariant);
      }
    }).then((): Promise<void> => {
      // Load initial data
      return this._loadLists();
    }).then((): Promise<void> => {
      if (this.properties.selectedList) {
        return this._loadColumns().then((): Promise<void> => {
          return this._loadFaqItems();
        });
      }
      return Promise.resolve();
    }).catch((error: Error) => {
      console.error('Error during initialization:', error);
      return Promise.resolve();
    });
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
    if (!this.domElement) {
      return;
    }

    try {
      // Set semantic colors as CSS variables
      if (theme.semanticColors) {
        const semanticColors: { [key: string]: string } = theme.semanticColors as { [key: string]: string };
        const semanticKeys = Object.keys(semanticColors);
        for (let i = 0; i < semanticKeys.length; i++) {
          const key = semanticKeys[i];
          const value = semanticColors[key];
          if (value) {
            this.domElement.style.setProperty('--' + key, value);
          }
        }
      }

      // Set palette colors as CSS variables
      if (theme.palette) {
        const palette: { [key: string]: string } = theme.palette as { [key: string]: string };
        const paletteKeys = Object.keys(palette);
        for (let i = 0; i < paletteKeys.length; i++) {
          const key = paletteKeys[i];
          const value = palette[key];
          if (value) {
            this.domElement.style.setProperty('--' + key, value);
          }
        }
      }
    } catch (error) {
      console.error('Error setting CSS variables:', error);
    }
  }

  /**
   * Extract unique categories from FAQ items
   */
  private _extractCategories(): void {
    const categorySet: { [key: string]: boolean } = {};
    this._categories = [];

    for (let i = 0; i < this._faqItems.length; i++) {
      const category = this._faqItems[i].category;
      if (category && !categorySet[category]) {
        categorySet[category] = true;
        this._categories.push(category);
      }
    }

    // Sort categories alphabetically
    this._categories.sort();
  }

  /**
   * Get filtered FAQ items based on selected category
   */
  private _getFilteredItems(): IFaqItem[] {
    if (this._selectedCategory === 'All') {
      return this._faqItems;
    }

    const filtered: IFaqItem[] = [];
    for (let i = 0; i < this._faqItems.length; i++) {
      if (this._faqItems[i].category === this._selectedCategory) {
        filtered.push(this._faqItems[i]);
      }
    }
    return filtered;
  }

  /**
   * Get font size value with px suffix
   */
  private _getFontSize(size: string | undefined, defaultSize: string): string {
    return (size || defaultSize) + 'px';
  }

  /**
   * Render the web part
   */
  public render(): void {
    try {
      // Build inline styles from theme for elements that need direct styling
      const semanticColors = this._themeVariant ? this._themeVariant.semanticColors : undefined;
      const palette = this._themeVariant ? this._themeVariant.palette : undefined;

      // Header background using theme primary color
      const headerBgColor = (palette && palette.themePrimary) ? palette.themePrimary : '#0078d4';
      const headerBgColorDark = (palette && palette.themeDark) ? palette.themeDark : '#005a9e';

      // Body colors
      const bodyBackground = (semanticColors && semanticColors.bodyBackground) ? semanticColors.bodyBackground : '#ffffff';
      const bodyText = (semanticColors && semanticColors.bodyText) ? semanticColors.bodyText : '#323130';
      const bodySubtext = (semanticColors && semanticColors.bodySubtext) ? semanticColors.bodySubtext : '#605e5c';

      // Interactive colors
      const linkColor = (semanticColors && semanticColors.link) ? semanticColors.link : ((palette && palette.themePrimary) ? palette.themePrimary : '#0078d4');

      // Border and divider colors
      const bodyDivider = (semanticColors && semanticColors.bodyDivider) ? semanticColors.bodyDivider : '#edebe9';

      // Hover states
      const listItemBackgroundHovered = (semanticColors && semanticColors.listItemBackgroundHovered) ? semanticColors.listItemBackgroundHovered : '#f3f2f1';

      // Card/container background
      const cardBackground = (semanticColors && semanticColors.cardStandoutBackground) ? semanticColors.cardStandoutBackground : bodyBackground;

      // Attachment area background
      const neutralLighter = (palette && palette.neutralLighter) ? palette.neutralLighter : '#f4f4f4';
      const neutralTertiary = (palette && palette.neutralTertiary) ? palette.neutralTertiary : '#a6a6a6';
      const neutralLight = (palette && palette.neutralLight) ? palette.neutralLight : '#eaeaea';

      // Font sizes with defaults
      const webpartTitleFontSize = this._getFontSize(this.properties.webpartTitleFontSize, '24');
      const webpartDescriptionFontSize = this._getFontSize(this.properties.webpartDescriptionFontSize, '14');
      const faqTitleFontSize = this._getFontSize(this.properties.faqTitleFontSize, '16');
      const faqDescriptionFontSize = this._getFontSize(this.properties.faqDescriptionFontSize, '14');

      let faqItemsHtml = '';
      let categoryTabsHtml = '';

      if (!this.properties.selectedList) {
        faqItemsHtml = '<div class="' + styles.emptyState + '" style="color: ' + bodySubtext + ';">' +
          '<svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="' + neutralTertiary + '" stroke-width="1.5">' +
          '<circle cx="12" cy="12" r="10"/>' +
          '<path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"/>' +
          '<line x1="12" y1="17" x2="12.01" y2="17"/>' +
          '</svg>' +
          '<p>Please configure the web part by selecting a list from the property pane.</p>' +
          '</div>';
      } else {
        // Build category tabs if category column is selected
        if (this.properties.categoryColumn && this._categories.length > 0) {
          const tabsArray: string[] = [];
          
          // Add "All" tab
          const allActiveClass = this._selectedCategory === 'All' ? ' ' + styles.activeTab : '';
          tabsArray.push(
            '<button class="' + styles.categoryTab + allActiveClass + '" ' +
            'data-category="All" ' +
            'style="color: ' + (this._selectedCategory === 'All' ? headerBgColor : bodyText) + '; ' +
            'border-bottom-color: ' + (this._selectedCategory === 'All' ? headerBgColor : 'transparent') + ';">' +
            'All' +
            '</button>'
          );

          // Add category tabs
          for (let i = 0; i < this._categories.length; i++) {
            const category = this._categories[i];
            const isActive = this._selectedCategory === category;
            const activeClass = isActive ? ' ' + styles.activeTab : '';
            tabsArray.push(
              '<button class="' + styles.categoryTab + activeClass + '" ' +
              'data-category="' + this._escapeHtml(category) + '" ' +
              'style="color: ' + (isActive ? headerBgColor : bodyText) + '; ' +
              'border-bottom-color: ' + (isActive ? headerBgColor : 'transparent') + ';">' +
              this._escapeHtml(category) +
              '</button>'
            );
          }

          categoryTabsHtml = '<div class="' + styles.categoryTabs + '" style="border-bottom-color: ' + neutralLight + ';">' +
            tabsArray.join('') +
            '</div>';
        }

        // Get filtered items
        const filteredItems = this._getFilteredItems();

        if (filteredItems.length === 0) {
          faqItemsHtml = '<div class="' + styles.emptyState + '" style="color: ' + bodySubtext + ';">' +
            '<svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="' + neutralTertiary + '" stroke-width="1.5">' +
            '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>' +
            '<polyline points="14 2 14 8 20 8"/>' +
            '<line x1="12" y1="18" x2="12" y2="12"/>' +
            '<line x1="9" y1="15" x2="15" y2="15"/>' +
            '</svg>' +
            '<p>No FAQ items found' + (this._selectedCategory !== 'All' ? ' in this category' : '') + '.</p>' +
            '</div>';
        } else {
          const itemsHtmlArray: string[] = [];
          for (let index = 0; index < filteredItems.length; index++) {
            const item = filteredItems[index];
            let attachmentsHtml = '';

            if (item.attachments && item.attachments.length > 0) {
              const attachmentLinksArray: string[] = [];
              for (let j = 0; j < item.attachments.length; j++) {
                const att = item.attachments[j];
                attachmentLinksArray.push(
                  '<a href="' + att.url + '" target="_blank" rel="noopener noreferrer" class="' + styles.attachmentLink + '" style="color: ' + linkColor + ';">' +
                  '<span class="' + styles.attachmentIcon + '">' +
                  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
                  '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>' +
                  '<polyline points="14 2 14 8 20 8"/>' +
                  '</svg>' +
                  '</span>' +
                  this._escapeHtml(att.fileName) +
                  '</a>'
                );
              }

              attachmentsHtml = '<div class="' + styles.attachments + '" style="background-color: ' + neutralLighter + ';">' +
                '<div class="' + styles.attachmentsLabel + '" style="color: ' + neutralTertiary + ';">' +
                '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
                '<path d="M21.44 11.05l-9.19 9.19a6 6 0 0 1-8.49-8.49l9.19-9.19a4 4 0 0 1 5.66 5.66l-9.2 9.19a2 2 0 0 1-2.83-2.83l8.49-8.48"/>' +
                '</svg>' +
                ' Attachments' +
                '</div>' +
                attachmentLinksArray.join('') +
                '</div>';
            }

            itemsHtmlArray.push(
              '<div class="' + styles.faqItem + '" data-index="' + index + '">' +
              '<div class="' + styles.faqQuestion + '" style="border-bottom-color: ' + bodyDivider + ';" data-hover-bg="' + listItemBackgroundHovered + '">' +
              '<span class="' + styles.faqQuestionText + '" style="color: ' + bodyText + '; font-size: ' + faqTitleFontSize + ';">' +
              this._escapeHtml(item.title) +
              '</span>' +
              '<span class="' + styles.faqChevron + '" style="color: ' + bodySubtext + ';">' +
              '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
              '<path d="M6 9l6 6 6-6"/>' +
              '</svg>' +
              '</span>' +
              '</div>' +
              '<div class="' + styles.faqAnswer + '">' +
              '<div class="' + styles.faqAnswerContent + '" style="color: ' + bodySubtext + '; font-size: ' + faqDescriptionFontSize + ';">' +
              this._formatDescription(item.description) +
              '</div>' +
              attachmentsHtml +
              '</div>' +
              '</div>'
            );
          }
          faqItemsHtml = itemsHtmlArray.join('');
        }
      }

      // Build header section
      let headerHtml = '';
      if (this.properties.webpartTitle || this.properties.webpartDescription) {
        headerHtml = '<div class="' + styles.faqHeader + '" style="background: linear-gradient(135deg, ' + headerBgColor + ' 0%, ' + headerBgColorDark + ' 100%);">';
        if (this.properties.webpartTitle) {
          headerHtml += '<h2 class="' + styles.faqTitle + '" style="font-size: ' + webpartTitleFontSize + ';">' + this._escapeHtml(this.properties.webpartTitle) + '</h2>';
        }
        if (this.properties.webpartDescription) {
          headerHtml += '<p class="' + styles.faqDescription + '" style="font-size: ' + webpartDescriptionFontSize + ';">' + this._escapeHtml(this.properties.webpartDescription) + '</p>';
        }
        headerHtml += '</div>';
      }

      this.domElement.innerHTML = '<div class="' + styles.customFaq + '" style="background-color: ' + cardBackground + ';">' +
        headerHtml +
        categoryTabsHtml +
        '<div class="' + styles.faqList + '">' +
        faqItemsHtml +
        '</div>' +
        '</div>';

      // Attach event listeners for accordion functionality
      this._attachEventListeners();
      // Attach event listeners for category tabs
      this._attachTabEventListeners();
    } catch (error) {
      console.error('Error during render:', error);
      this.domElement.innerHTML = '<div style="padding: 20px; color: red;">Error rendering web part. Please check console for details.</div>';
    }
  }

  /**
   * Attach click event listeners to category tabs
   */
  private _attachTabEventListeners(): void {
    const tabs = this.domElement.querySelectorAll('.' + styles.categoryTab);
    const self = this;

    for (let i = 0; i < tabs.length; i++) {
      const tab = tabs[i] as HTMLElement;
      tab.addEventListener('click', function(): void {
        const category = tab.getAttribute('data-category');
        if (category) {
          self._selectedCategory = category;
          self.render();
        }
      });
    }
  }

  /**
   * Attach click event listeners to FAQ items
   */
  private _attachEventListeners(): void {
    const faqItems = this.domElement.querySelectorAll('.' + styles.faqItem);
    const questions = this.domElement.querySelectorAll('.' + styles.faqQuestion);
    const self = this;

    for (let index = 0; index < questions.length; index++) {
      const questionElement = questions[index] as HTMLElement;

      // Add hover effect
      questionElement.addEventListener('mouseenter', function(): void {
        const hoverBg = questionElement.getAttribute('data-hover-bg');
        if (hoverBg) {
          questionElement.style.backgroundColor = hoverBg;
        }
      });

      questionElement.addEventListener('mouseleave', function(): void {
        questionElement.style.backgroundColor = '';
      });

      // Add click handler for accordion - use closure to capture index
      (function(idx: number): void {
        questionElement.addEventListener('click', function(): void {
          const faqItem = faqItems[idx];

          if (!self.properties.allowMultipleExpanded) {
            // Close all other items
            for (let i = 0; i < faqItems.length; i++) {
              if (i !== idx && faqItems[i].classList.contains(styles.expanded)) {
                faqItems[i].classList.remove(styles.expanded);
              }
            }
          }

          // Toggle current item
          if (faqItem.classList.contains(styles.expanded)) {
            faqItem.classList.remove(styles.expanded);
          } else {
            faqItem.classList.add(styles.expanded);
          }
        });
      })(index);
    }
  }

  /**
   * Escape HTML to prevent XSS
   */
  private _escapeHtml(text: string): string {
    if (!text) {
      return '';
    }
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  /**
   * Format description text (handle rich text or plain text)
   */
  private _formatDescription(description: string): string {
    if (!description) {
      return '';
    }

    // Check if it's HTML content (from multi-line rich text field)
    if (description.indexOf('<') !== -1 && description.indexOf('>') !== -1) {
      // Return as-is for HTML content
      return description;
    }

    // Plain text - convert line breaks to <br>
    return this._escapeHtml(description).replace(/\n/g, '<br>');
  }

  /**
   * Load all lists from the current site
   */
  private _loadLists(): Promise<void> {
    return this._spService.getLists()
      .then((lists: IListInfo[]) => {
        this._lists = lists;
      })
      .catch((error: Error) => {
        console.error('Error loading lists:', error);
        this._lists = [];
      });
  }

  /**
   * Load columns for the selected list
   */
  private _loadColumns(): Promise<void> {
    if (!this.properties.selectedList) {
      this._columns = [];
      return Promise.resolve();
    }

    return this._spService.getListColumns(this.properties.selectedList)
      .then((columns: IColumnInfo[]) => {
        this._columns = columns;
      })
      .catch((error: Error) => {
        console.error('Error loading columns:', error);
        this._columns = [];
      });
  }

  /**
   * Load FAQ items from the selected list
   */
  private _loadFaqItems(): Promise<void> {
    if (!this.properties.selectedList || !this.properties.titleColumn || !this.properties.descriptionColumn) {
      this._faqItems = [];
      return Promise.resolve();
    }

    return this._spService.getListItems(
      this.properties.selectedList,
      this.properties.titleColumn,
      this.properties.descriptionColumn,
      this.properties.categoryColumn || undefined
    )
      .then((items: IFaqItem[]) => {
        this._faqItems = items;
        this._extractCategories();
        // Reset to All if current category no longer exists
        let categoryExists = false;
        for (let i = 0; i < this._categories.length; i++) {
          if (this._categories[i] === this._selectedCategory) {
            categoryExists = true;
            break;
          }
        }
        if (this._selectedCategory !== 'All' && !categoryExists) {
          this._selectedCategory = 'All';
        }
      })
      .catch((error: Error) => {
        console.error('Error loading FAQ items:', error);
        this._faqItems = [];
        this._categories = [];
      });
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
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: string | boolean, newValue: string | boolean): void {
    const self = this;

    if (propertyPath === 'selectedList' && newValue !== oldValue) {
      // Reset column selections when list changes
      this.properties.titleColumn = '';
      this.properties.descriptionColumn = '';
      this.properties.categoryColumn = '';
      this._faqItems = [];
      this._categories = [];
      this._selectedCategory = 'All';

      // Load columns for new list
      this._loadColumns().then(function(): void {
        // Refresh property pane to show new columns
        self.context.propertyPane.refresh();
        self.render();
      });
    } else if ((propertyPath === 'titleColumn' || propertyPath === 'descriptionColumn' || propertyPath === 'categoryColumn') && newValue !== oldValue) {
      // Reload FAQ items when columns change
      this._loadFaqItems().then(function(): void {
        self.render();
      });
    } else {
      this.render();
    }
  }

  /**
   * Get dropdown options for lists
   */
  private _getListOptions(): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = [
      { key: '', text: '-- Select a list --' }
    ];

    for (let i = 0; i < this._lists.length; i++) {
      options.push({ key: this._lists[i].id, text: this._lists[i].title });
    }

    return options;
  }

  /**
   * Get dropdown options for title columns (text columns only)
   */
  private _getTitleColumnOptions(): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = [
      { key: '', text: '-- Select a column --' }
    ];

    for (let i = 0; i < this._columns.length; i++) {
      const col = this._columns[i];
      if (col.type === 'Text' || col.type === 'Note') {
        options.push({ key: col.internalName, text: col.title });
      }
    }

    return options;
  }

  /**
   * Get dropdown options for description columns (text and multi-line text columns)
   */
  private _getDescriptionColumnOptions(): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = [
      { key: '', text: '-- Select a column --' }
    ];

    for (let i = 0; i < this._columns.length; i++) {
      const col = this._columns[i];
      if (col.type === 'Text' || col.type === 'Note') {
        const typeLabel = col.type === 'Note' ? 'Multi-line' : 'Single-line';
        options.push({ key: col.internalName, text: col.title + ' (' + typeLabel + ')' });
      }
    }

    return options;
  }

  /**
   * Get dropdown options for category column (text and choice columns)
   */
  private _getCategoryColumnOptions(): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = [
      { key: '', text: '-- No category (disable tabs) --' }
    ];

    for (let i = 0; i < this._columns.length; i++) {
      const col = this._columns[i];
      if (col.type === 'Text' || col.type === 'Choice') {
        options.push({ key: col.internalName, text: col.title });
      }
    }

    return options;
  }

  /**
   * Get font size options
   */
  private _getFontSizeOptions(): IPropertyPaneDropdownOption[] {
    return [
      { key: '10', text: '10px - Extra Small' },
      { key: '12', text: '12px - Small' },
      { key: '14', text: '14px - Normal' },
      { key: '16', text: '16px - Medium' },
      { key: '18', text: '18px - Large' },
      { key: '20', text: '20px - Extra Large' },
      { key: '24', text: '24px - Heading' },
      { key: '28', text: '28px - Large Heading' },
      { key: '32', text: '32px - Extra Large Heading' }
    ];
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
                  placeholder: 'Enter a title for the FAQ section'
                }),
                PropertyPaneDropdown('webpartTitleFontSize', {
                  label: 'Webpart Title Font Size',
                  options: this._getFontSizeOptions(),
                  selectedKey: this.properties.webpartTitleFontSize || '24'
                }),
                PropertyPaneTextField('webpartDescription', {
                  label: 'Webpart Description',
                  placeholder: 'Enter a description',
                  multiline: true,
                  rows: 3
                }),
                PropertyPaneDropdown('webpartDescriptionFontSize', {
                  label: 'Webpart Description Font Size',
                  options: this._getFontSizeOptions(),
                  selectedKey: this.properties.webpartDescriptionFontSize || '14'
                })
              ]
            },
            {
              groupName: 'FAQ Item Settings',
              groupFields: [
                PropertyPaneDropdown('faqTitleFontSize', {
                  label: 'FAQ Question Font Size',
                  options: this._getFontSizeOptions(),
                  selectedKey: this.properties.faqTitleFontSize || '16'
                }),
                PropertyPaneDropdown('faqDescriptionFontSize', {
                  label: 'FAQ Answer Font Size',
                  options: this._getFontSizeOptions(),
                  selectedKey: this.properties.faqDescriptionFontSize || '14'
                })
              ]
            },
            {
              groupName: 'Data Source',
              groupFields: [
                PropertyPaneDropdown('selectedList', {
                  label: 'Select List',
                  options: this._getListOptions()
                }),
                PropertyPaneDropdown('titleColumn', {
                  label: 'Title Column',
                  options: this._getTitleColumnOptions(),
                  disabled: !this.properties.selectedList
                }),
                PropertyPaneDropdown('descriptionColumn', {
                  label: 'Description Column',
                  options: this._getDescriptionColumnOptions(),
                  disabled: !this.properties.selectedList
                }),
                PropertyPaneDropdown('categoryColumn', {
                  label: 'Category Column (for tabs)',
                  options: this._getCategoryColumnOptions(),
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
                  offText: 'No'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}