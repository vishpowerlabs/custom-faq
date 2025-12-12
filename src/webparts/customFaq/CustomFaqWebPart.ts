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
  enableSearch: boolean;
  searchInHeader: boolean;
  highlightSearchMatches: boolean;
  searchInAnswers: boolean;
  showResultsCount: boolean;
  searchPlaceholder: string;
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
  private _searchQuery: string = '';

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
          this._themeVariant = this._themeProvider.tryGetTheme();
          this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);
        }
      } catch (error) {
        console.log('ThemeProvider not available:', error);
      }

      if (this._themeVariant) {
        this._setCSSVariables(this._themeVariant);
      }
    }).then((): Promise<void> => {
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

  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    if (this._themeVariant) {
      this._setCSSVariables(this._themeVariant);
    }
    this.render();
  }

  private _setCSSVariables(theme: IReadonlyTheme): void {
    if (!this.domElement) {
      return;
    }

    try {
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

    this._categories.sort();
  }

  /**
   * Get filtered FAQ items based on selected category and search query
   */
  private _getFilteredItems(): IFaqItem[] {
    let filtered: IFaqItem[] = [];

    // First filter by category
    if (this._selectedCategory === 'All') {
      filtered = this._faqItems.slice();
    } else {
      for (let i = 0; i < this._faqItems.length; i++) {
        if (this._faqItems[i].category === this._selectedCategory) {
          filtered.push(this._faqItems[i]);
        }
      }
    }

    // Then filter by search query
    if (this._searchQuery && this._searchQuery.trim() !== '') {
      const query = this._searchQuery.toLowerCase().trim();
      const searchResults: IFaqItem[] = [];

      for (let i = 0; i < filtered.length; i++) {
        const item = filtered[i];
        const titleMatch = item.title && item.title.toLowerCase().indexOf(query) !== -1;
        let answerMatch = false;

        if (this.properties.searchInAnswers && item.description) {
          // Strip HTML tags for search using loop-based sanitization
          const plainDescription = this._stripHtmlTags(item.description);
          answerMatch = plainDescription.toLowerCase().indexOf(query) !== -1;
        }

        if (titleMatch || answerMatch) {
          searchResults.push(item);
        }
      }

      filtered = searchResults;
    }

    return filtered;
  }

  /**
   * Get total count of items (before search filter, after category filter)
   */
  private _getTotalItemsInCategory(): number {
    if (this._selectedCategory === 'All') {
      return this._faqItems.length;
    }

    let count = 0;
    for (let i = 0; i < this._faqItems.length; i++) {
      if (this._faqItems[i].category === this._selectedCategory) {
        count++;
      }
    }
    return count;
  }

  /**
   * Highlight search matches in text
   */
  private _highlightText(text: string, isHtml: boolean): string {
    if (!this._searchQuery || this._searchQuery.trim() === '' || !this.properties.highlightSearchMatches) {
      return text;
    }

    const query = this._searchQuery.trim();
    
    if (isHtml) {
      // For HTML content, use DOM-based approach for safe highlighting
      return this._highlightInHtml(text, query);
    } else {
      // For plain text, simple replacement is safe
      const regex = new RegExp('(' + this._escapeRegExp(query) + ')', 'gi');
      return text.replace(regex, '<mark class="' + styles.searchHighlight + '">$1</mark>');
    }
  }

  /**
   * Safely highlight text within HTML content using DOM parsing
   */
  private _highlightInHtml(html: string, query: string): string {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString('<div>' + html + '</div>', 'text/html');
      const container = doc.body.firstChild as HTMLElement;
      
      if (!container) {
        return html;
      }

      // Walk through text nodes only
      const walker = doc.createTreeWalker(
        container,
        NodeFilter.SHOW_TEXT,
        null
      );

      const textNodes: Text[] = [];
      let node: Text | null;
      while ((node = walker.nextNode() as Text | null)) {
        textNodes.push(node);
      }

      const lowerQuery = query.toLowerCase();
      
      for (let i = 0; i < textNodes.length; i++) {
        const textNode = textNodes[i];
        const text = textNode.textContent || '';
        const lowerText = text.toLowerCase();
        const index = lowerText.indexOf(lowerQuery);
        
        if (index !== -1 && textNode.parentNode) {
          const before = text.substring(0, index);
          const match = text.substring(index, index + query.length);
          const after = text.substring(index + query.length);
          
          const fragment = doc.createDocumentFragment();
          
          if (before) {
            fragment.appendChild(doc.createTextNode(before));
          }
          
          const mark = doc.createElement('mark');
          mark.className = styles.searchHighlight;
          mark.textContent = match;
          fragment.appendChild(mark);
          
          if (after) {
            fragment.appendChild(doc.createTextNode(after));
          }
          
          textNode.parentNode.replaceChild(fragment, textNode);
        }
      }

      return container.innerHTML;
    } catch (error) {
      // Fallback: return original HTML if parsing fails
      console.warn('HTML highlighting failed:', error);
      return html;
    }
  }

  /**
   * Securely strip HTML tags using DOM-based parsing
   * Addresses CodeQL js/incomplete-multi-character-sanitization
   * and js/redos (ReDoS vulnerability)
   */
  private _stripHtmlTags(input: string): string {
    if (!input) {
      return '';
    }

    try {
      // Use DOMParser for safe HTML stripping - no regex vulnerability
      const parser = new DOMParser();
      const doc = parser.parseFromString('<div>' + input + '</div>', 'text/html');
      const container = doc.body.firstChild as HTMLElement;
      
      if (!container) {
        return this._fallbackStripTags(input);
      }

      // Get text content only (strips all HTML)
      let result = container.textContent || '';

      // Decode common HTML entities
      result = this._decodeHtmlEntities(result);

      return result;
    } catch (error) {
      // Fallback if DOMParser fails
      console.warn('DOMParser failed, using fallback:', error);
      return this._fallbackStripTags(input);
    }
  }

  /**
   * Fallback tag stripper using character-by-character parsing
   * Safe from ReDoS as it doesn't use regex with backtracking
   */
  private _fallbackStripTags(input: string): string {
    let result = '';
    let inTag = false;
    const chars = input.split('');
    let idx = 0;
    const len = chars.length;

    while (idx < len) {
      const char = chars[idx];
      if (char === '<') {
        inTag = true;
      } else if (char === '>') {
        inTag = false;
      } else if (!inTag) {
        result += char;
      }
      idx++;
    }

    return this._decodeHtmlEntities(result);
  }

  /**
   * Decode common HTML entities
   */
  private _decodeHtmlEntities(text: string): string {
    const entities: { [key: string]: string } = {
      '&nbsp;': ' ',
      '&amp;': '&',
      '&lt;': '<',
      '&gt;': '>',
      '&quot;': '"',
      '&#39;': "'",
      '&#x27;': "'",
      '&apos;': "'"
    };

    let result = text;
    const keys = Object.keys(entities);
    for (let i = 0; i < keys.length; i++) {
      const entity = keys[i];
      // Use split/join instead of regex to avoid ReDoS
      result = result.split(entity).join(entities[entity]);
    }

    return result;
  }

  /**
   * Escape special regex characters
   */
  private _escapeRegExp(text: string): string {
    return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  private _getFontSize(size: string | undefined, defaultSize: string): string {
    return (size || defaultSize) + 'px';
  }

  /**
   * Render the web part
   */
  public render(): void {
    try {
      const semanticColors = this._themeVariant ? this._themeVariant.semanticColors : undefined;
      const palette = this._themeVariant ? this._themeVariant.palette : undefined;

      const headerBgColor = (palette && palette.themePrimary) ? palette.themePrimary : '#0078d4';
      const headerBgColorDark = (palette && palette.themeDark) ? palette.themeDark : '#005a9e';
      const bodyBackground = (semanticColors && semanticColors.bodyBackground) ? semanticColors.bodyBackground : '#ffffff';
      const bodyText = (semanticColors && semanticColors.bodyText) ? semanticColors.bodyText : '#323130';
      const bodySubtext = (semanticColors && semanticColors.bodySubtext) ? semanticColors.bodySubtext : '#605e5c';
      const linkColor = (semanticColors && semanticColors.link) ? semanticColors.link : ((palette && palette.themePrimary) ? palette.themePrimary : '#0078d4');
      const bodyDivider = (semanticColors && semanticColors.bodyDivider) ? semanticColors.bodyDivider : '#edebe9';
      const listItemBackgroundHovered = (semanticColors && semanticColors.listItemBackgroundHovered) ? semanticColors.listItemBackgroundHovered : '#f3f2f1';
      const cardBackground = (semanticColors && semanticColors.cardStandoutBackground) ? semanticColors.cardStandoutBackground : bodyBackground;
      const neutralLighter = (palette && palette.neutralLighter) ? palette.neutralLighter : '#f4f4f4';
      const neutralTertiary = (palette && palette.neutralTertiary) ? palette.neutralTertiary : '#a6a6a6';
      const neutralLight = (palette && palette.neutralLight) ? palette.neutralLight : '#eaeaea';
      const inputBackground = (semanticColors && semanticColors.inputBackground) ? semanticColors.inputBackground : '#ffffff';
      const inputBorder = (semanticColors && semanticColors.inputBorder) ? semanticColors.inputBorder : '#8a8886';

      const webpartTitleFontSize = this._getFontSize(this.properties.webpartTitleFontSize, '24');
      const webpartDescriptionFontSize = this._getFontSize(this.properties.webpartDescriptionFontSize, '14');
      const faqTitleFontSize = this._getFontSize(this.properties.faqTitleFontSize, '16');
      const faqDescriptionFontSize = this._getFontSize(this.properties.faqDescriptionFontSize, '14');

      const searchPlaceholder = this.properties.searchPlaceholder || 'Search FAQs...';

      let faqItemsHtml = '';
      let categoryTabsHtml = '';
      let searchHtml = '';
      let searchInHeaderHtml = '';
      let resultsInfoHtml = '';

      // Build search input HTML
      const searchInputHtml = '<div class="' + styles.searchContainer + '">' +
        '<svg class="' + styles.searchIcon + '" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
        '<circle cx="11" cy="11" r="8"/>' +
        '<path d="M21 21l-4.35-4.35"/>' +
        '</svg>' +
        '<input type="text" class="' + styles.searchInput + '" ' +
        'placeholder="' + this._escapeHtml(searchPlaceholder) + '" ' +
        'value="' + this._escapeHtml(this._searchQuery) + '" ' +
        'style="background-color: ' + inputBackground + '; border-color: ' + inputBorder + '; color: ' + bodyText + ';">' +
        (this._searchQuery ? '<button class="' + styles.clearButton + '" title="Clear search" style="color: ' + bodySubtext + ';">' +
        '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
        '<line x1="18" y1="6" x2="6" y2="18"/>' +
        '<line x1="6" y1="6" x2="18" y2="18"/>' +
        '</svg>' +
        '</button>' : '') +
        '</div>';

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
        // Build search section (below header)
        if (this.properties.enableSearch && !this.properties.searchInHeader) {
          searchHtml = '<div class="' + styles.searchSection + '" style="background-color: ' + neutralLighter + '; border-bottom-color: ' + bodyDivider + ';">' +
            searchInputHtml +
            '</div>';
        }

        // Build search in header
        if (this.properties.enableSearch && this.properties.searchInHeader) {
          searchInHeaderHtml = '<div class="' + styles.searchInHeader + '">' + searchInputHtml + '</div>';
        }

        // Build category tabs
        if (this.properties.categoryColumn && this._categories.length > 0) {
          const tabsArray: string[] = [];
          
          const allActiveClass = this._selectedCategory === 'All' ? ' ' + styles.activeTab : '';
          tabsArray.push(
            '<button class="' + styles.categoryTab + allActiveClass + '" ' +
            'data-category="All" ' +
            'style="color: ' + (this._selectedCategory === 'All' ? headerBgColor : bodyText) + '; ' +
            'border-bottom-color: ' + (this._selectedCategory === 'All' ? headerBgColor : 'transparent') + ';">' +
            'All' +
            '</button>'
          );

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
        const totalItems = this._getTotalItemsInCategory();

        // Build results info
        if (this.properties.enableSearch && this.properties.showResultsCount && this._searchQuery && this._searchQuery.trim() !== '') {
          resultsInfoHtml = '<div class="' + styles.resultsInfo + '" style="color: ' + bodySubtext + ';">' +
            'Showing <strong style="color: ' + bodyText + ';">' + filteredItems.length + '</strong> of ' +
            '<strong style="color: ' + bodyText + ';">' + totalItems + '</strong> FAQs matching ' +
            '"<strong style="color: ' + bodyText + ';">' + this._escapeHtml(this._searchQuery) + '</strong>"' +
            '</div>';
        }

        if (filteredItems.length === 0) {
          const noResultsMessage = this._searchQuery && this._searchQuery.trim() !== '' 
            ? 'No FAQs found matching "' + this._escapeHtml(this._searchQuery) + '". Try different keywords or clear the search.'
            : 'No FAQ items found' + (this._selectedCategory !== 'All' ? ' in this category' : '') + '.';

          faqItemsHtml = '<div class="' + styles.emptyState + '" style="color: ' + bodySubtext + ';">' +
            '<svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="' + neutralTertiary + '" stroke-width="1.5">' +
            (this._searchQuery ? 
              '<circle cx="11" cy="11" r="8"/><path d="M21 21l-4.35-4.35"/><line x1="8" y1="11" x2="14" y2="11"/>' :
              '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/>'
            ) +
            '</svg>' +
            '<p>' + noResultsMessage + '</p>' +
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

            // Apply highlighting to title
            const highlightedTitle = this._highlightText(this._escapeHtml(item.title), false);
            
            // Apply highlighting to description
            const formattedDescription = this._formatDescription(item.description);
            const highlightedDescription = this.properties.searchInAnswers 
              ? this._highlightText(formattedDescription, formattedDescription.indexOf('<') !== -1)
              : formattedDescription;

            itemsHtmlArray.push(
              '<div class="' + styles.faqItem + '" data-index="' + index + '">' +
              '<div class="' + styles.faqQuestion + '" style="border-bottom-color: ' + bodyDivider + ';" data-hover-bg="' + listItemBackgroundHovered + '">' +
              '<span class="' + styles.faqQuestionText + '" style="color: ' + bodyText + '; font-size: ' + faqTitleFontSize + ';">' +
              highlightedTitle +
              '</span>' +
              '<span class="' + styles.faqChevron + '" style="color: ' + bodySubtext + ';">' +
              '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
              '<path d="M6 9l6 6 6-6"/>' +
              '</svg>' +
              '</span>' +
              '</div>' +
              '<div class="' + styles.faqAnswer + '">' +
              '<div class="' + styles.faqAnswerContent + '" style="color: ' + bodySubtext + '; font-size: ' + faqDescriptionFontSize + ';">' +
              highlightedDescription +
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
      if (this.properties.webpartTitle || this.properties.webpartDescription || searchInHeaderHtml) {
        headerHtml = '<div class="' + styles.faqHeader + '" style="background: linear-gradient(135deg, ' + headerBgColor + ' 0%, ' + headerBgColorDark + ' 100%);">';
        if (this.properties.webpartTitle) {
          headerHtml += '<h2 class="' + styles.faqTitle + '" style="font-size: ' + webpartTitleFontSize + ';">' + this._escapeHtml(this.properties.webpartTitle) + '</h2>';
        }
        if (this.properties.webpartDescription) {
          headerHtml += '<p class="' + styles.faqDescription + '" style="font-size: ' + webpartDescriptionFontSize + ';">' + this._escapeHtml(this.properties.webpartDescription) + '</p>';
        }
        headerHtml += searchInHeaderHtml;
        headerHtml += '</div>';
      }

      this.domElement.innerHTML = '<div class="' + styles.customFaq + '" style="background-color: ' + cardBackground + ';">' +
        headerHtml +
        searchHtml +
        resultsInfoHtml +
        categoryTabsHtml +
        '<div class="' + styles.faqList + '">' +
        faqItemsHtml +
        '</div>' +
        '</div>';

      this._attachEventListeners();
      this._attachTabEventListeners();
      this._attachSearchEventListeners();
    } catch (error) {
      console.error('Error during render:', error);
      this.domElement.innerHTML = '<div style="padding: 20px; color: red;">Error rendering web part. Please check console for details.</div>';
    }
  }

  /**
   * Attach search event listeners
   */
  private _attachSearchEventListeners(): void {
    const searchInput = this.domElement.querySelector('.' + styles.searchInput) as HTMLInputElement;
    const clearButton = this.domElement.querySelector('.' + styles.clearButton) as HTMLButtonElement;
    const self = this;

    if (searchInput) {
      // Debounce search input
      let debounceTimer: number | undefined;
      
      searchInput.addEventListener('input', function(): void {
        if (debounceTimer) {
          clearTimeout(debounceTimer);
        }
        debounceTimer = setTimeout(function(): void {
          self._searchQuery = searchInput.value;
          self.render();
        }, 300) as unknown as number;
      });

      // Handle Enter key
      searchInput.addEventListener('keydown', function(e: KeyboardEvent): void {
        if (e.key === 'Escape') {
          self._searchQuery = '';
          self.render();
        }
      });

      // Focus the input if it had a value (maintain focus after re-render)
      if (this._searchQuery) {
        searchInput.focus();
        searchInput.setSelectionRange(searchInput.value.length, searchInput.value.length);
      }
    }

    if (clearButton) {
      clearButton.addEventListener('click', function(): void {
        self._searchQuery = '';
        self.render();
      });
    }
  }

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

  private _attachEventListeners(): void {
    const faqItems = this.domElement.querySelectorAll('.' + styles.faqItem);
    const questions = this.domElement.querySelectorAll('.' + styles.faqQuestion);
    const self = this;

    for (let index = 0; index < questions.length; index++) {
      const questionElement = questions[index] as HTMLElement;

      questionElement.addEventListener('mouseenter', function(): void {
        const hoverBg = questionElement.getAttribute('data-hover-bg');
        if (hoverBg) {
          questionElement.style.backgroundColor = hoverBg;
        }
      });

      questionElement.addEventListener('mouseleave', function(): void {
        questionElement.style.backgroundColor = '';
      });

      (function(idx: number): void {
        questionElement.addEventListener('click', function(): void {
          const faqItem = faqItems[idx];

          if (!self.properties.allowMultipleExpanded) {
            for (let i = 0; i < faqItems.length; i++) {
              if (i !== idx && faqItems[i].classList.contains(styles.expanded)) {
                faqItems[i].classList.remove(styles.expanded);
              }
            }
          }

          if (faqItem.classList.contains(styles.expanded)) {
            faqItem.classList.remove(styles.expanded);
          } else {
            faqItem.classList.add(styles.expanded);
          }
        });
      })(index);
    }
  }

  private _escapeHtml(text: string): string {
    if (!text) {
      return '';
    }
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  private _formatDescription(description: string): string {
    if (!description) {
      return '';
    }

    // Check if content appears to contain HTML
    const hasOpenTag = description.indexOf('<') !== -1;
    const hasCloseTag = description.indexOf('>') !== -1;
    
    if (hasOpenTag && hasCloseTag) {
      // Content has HTML, return as-is (will be rendered as HTML)
      return description;
    }

    // Plain text: escape HTML entities and convert newlines
    return this._escapeHtml(description).replace(/\n/g, '<br>');
  }

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

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: string | boolean, newValue: string | boolean): void {
    const self = this;

    if (propertyPath === 'selectedList' && newValue !== oldValue) {
      this.properties.titleColumn = '';
      this.properties.descriptionColumn = '';
      this.properties.categoryColumn = '';
      this._faqItems = [];
      this._categories = [];
      this._selectedCategory = 'All';
      this._searchQuery = '';

      this._loadColumns().then(function(): void {
        self.context.propertyPane.refresh();
        self.render();
      });
    } else if ((propertyPath === 'titleColumn' || propertyPath === 'descriptionColumn' || propertyPath === 'categoryColumn') && newValue !== oldValue) {
      this._loadFaqItems().then(function(): void {
        self.render();
      });
    } else {
      this.render();
    }
  }

  private _getListOptions(): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = [
      { key: '', text: '-- Select a list --' }
    ];

    for (let i = 0; i < this._lists.length; i++) {
      options.push({ key: this._lists[i].id, text: this._lists[i].title });
    }

    return options;
  }

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
              groupName: 'Search Settings',
              groupFields: [
                PropertyPaneToggle('enableSearch', {
                  label: 'Enable Search',
                  onText: 'Yes',
                  offText: 'No'
                }),
                PropertyPaneToggle('searchInHeader', {
                  label: 'Search in Header',
                  onText: 'Yes (in header)',
                  offText: 'No (below header)',
                  disabled: !this.properties.enableSearch
                }),
                PropertyPaneToggle('highlightSearchMatches', {
                  label: 'Highlight Search Matches',
                  onText: 'Yes',
                  offText: 'No',
                  disabled: !this.properties.enableSearch
                }),
                PropertyPaneToggle('searchInAnswers', {
                  label: 'Search in Answers',
                  onText: 'Yes (titles & answers)',
                  offText: 'No (titles only)',
                  disabled: !this.properties.enableSearch
                }),
                PropertyPaneToggle('showResultsCount', {
                  label: 'Show Results Count',
                  onText: 'Yes',
                  offText: 'No',
                  disabled: !this.properties.enableSearch
                }),
                PropertyPaneTextField('searchPlaceholder', {
                  label: 'Search Placeholder Text',
                  placeholder: 'Search FAQs...',
                  disabled: !this.properties.enableSearch
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