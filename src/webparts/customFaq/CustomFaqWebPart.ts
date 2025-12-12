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
  topFaqColumn: string;
  topFaqColor: string;
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

  private _spService!: SharePointService;
  private _lists: IListInfo[] = [];
  private _columns: IColumnInfo[] = [];
  private _faqItems: IFaqItem[] = [];
  private _themeProvider: ThemeProvider | undefined;
  private _themeVariant: IReadonlyTheme | undefined;
  private _selectedCategory: string = 'All';
  private _categories: string[] = [];
  private _searchQuery: string = '';
  private readonly _topFaqCategoryKey: string = '__TOP_FAQ__';

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
        let i = 0;
        while (i < semanticKeys.length) {
          const key = semanticKeys[i];
          const value = semanticColors[key];
          if (value) {
            this.domElement.style.setProperty('--' + key, value);
          }
          i++;
        }
      }

      if (theme.palette) {
        const palette: { [key: string]: string } = theme.palette as { [key: string]: string };
        const paletteKeys = Object.keys(palette);
        let j = 0;
        while (j < paletteKeys.length) {
          const key = paletteKeys[j];
          const value = palette[key];
          if (value) {
            this.domElement.style.setProperty('--' + key, value);
          }
          j++;
        }
      }
    } catch (error) {
      console.error('Error setting CSS variables:', error);
    }
  }

  private _extractCategories(): void {
    const categorySet: { [key: string]: boolean } = {};
    this._categories = [];

    let i = 0;
    while (i < this._faqItems.length) {
      const category = this._faqItems[i].category;
      if (category && !categorySet[category]) {
        categorySet[category] = true;
        this._categories.push(category);
      }
      i++;
    }

    this._categories.sort();
  }

  private _isTopFaqEnabled(): boolean {
    return !!(this.properties.topFaqColumn && this.properties.topFaqColumn.trim() !== '');
  }

  private _hasTopFaqItems(): boolean {
    let i = 0;
    while (i < this._faqItems.length) {
      if (this._faqItems[i].isTopFaq) {
        return true;
      }
      i++;
    }
    return false;
  }

  private _ensureValidCategorySelection(): void {
    const validCategories: { [key: string]: boolean } = {
      'All': true
    };

    if (this._isTopFaqEnabled()) {
      validCategories[this._topFaqCategoryKey] = true;
    }

    let i = 0;
    while (i < this._categories.length) {
      validCategories[this._categories[i]] = true;
      i++;
    }

    if (!validCategories[this._selectedCategory]) {
      this._selectedCategory = 'All';
    }
  }

  private _sanitizeColorInput(color: string | undefined): string {
    if (!color) {
      return '';
    }

    const trimmed = color.trim();
    if (trimmed === '') {
      return '';
    }

    const hexRegex = /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/;
    const rgbRegex = /^rgba?\(\s*\d{1,3}\s*,\s*\d{1,3}\s*,\s*\d{1,3}(?:\s*,\s*(0|0?\.\d+|1(\.0)?))?\s*\)$/;
    const cssNameRegex = /^[a-zA-Z]+$/;

    if (hexRegex.test(trimmed) || rgbRegex.test(trimmed) || cssNameRegex.test(trimmed)) {
      return trimmed;
    }

    return '';
  }

  private _getTopFaqColor(fallbackColor: string): string {
    const sanitized = this._sanitizeColorInput(this.properties.topFaqColor);
    if (sanitized) {
      return sanitized;
    }
    return fallbackColor;
  }

  private _getTopFaqBadgePalette(color: string): { badgeBackground: string; badgeTextColor: string; badgeBorder: string; badgeShadow: string } {
    const badgeBackground = color;
    const badgeBorder = this._hexToRgba(color, 0.35) || 'rgba(0, 0, 0, 0.3)';
    const badgeShadow = this._hexToRgba(color, 0.45) || 'rgba(0, 0, 0, 0.25)';
    const badgeTextColor = this._getContrastingTextColor(color);

    return {
      badgeBackground,
      badgeTextColor,
      badgeBorder,
      badgeShadow
    };
  }

  private _hexToRgba(color: string, alpha: number): string | null {
    if (!color) {
      return null;
    }
    const trimmed = color.trim();
    if (!trimmed || trimmed.charAt(0) !== '#') {
      return null;
    }

    let hex = trimmed.substring(1);
    if (hex.length === 3) {
      hex = hex.charAt(0) + hex.charAt(0) + hex.charAt(1) + hex.charAt(1) + hex.charAt(2) + hex.charAt(2);
    }

    if (hex.length !== 6 || hex.match(/[^0-9a-fA-F]/)) {
      return null;
    }

    const r = parseInt(hex.substring(0, 2), 16);
    const g = parseInt(hex.substring(2, 4), 16);
    const b = parseInt(hex.substring(4, 6), 16);
    const normalizedAlpha = Math.max(0, Math.min(1, alpha));

    return 'rgba(' + r + ', ' + g + ', ' + b + ', ' + normalizedAlpha + ')';
  }

  private _parseHexColor(color: string): { r: number; g: number; b: number } | null {
    if (!color || color.charAt(0) !== '#') {
      return null;
    }

    let hex = color.substring(1);
    if (hex.length === 3) {
      hex = hex.charAt(0) + hex.charAt(0) + hex.charAt(1) + hex.charAt(1) + hex.charAt(2) + hex.charAt(2);
    }

    if (hex.length !== 6 || hex.match(/[^0-9a-fA-F]/)) {
      return null;
    }

    return {
      r: parseInt(hex.substring(0, 2), 16),
      g: parseInt(hex.substring(2, 4), 16),
      b: parseInt(hex.substring(4, 6), 16)
    };
  }

  private _relativeLuminance(r: number, g: number, b: number): number {
    const srgb = [r, g, b].map((value: number) => {
      const normalized = value / 255;
      return normalized <= 0.03928 ? normalized / 12.92 : Math.pow((normalized + 0.055) / 1.055, 2.4);
    });

    return 0.2126 * srgb[0] + 0.7152 * srgb[1] + 0.0722 * srgb[2];
  }

  private _getContrastingTextColor(color: string): string {
    const rgb = this._parseHexColor(color);
    if (!rgb) {
      return '#ffffff';
    }

    const luminance = this._relativeLuminance(rgb.r, rgb.g, rgb.b);
    return luminance > 0.6 ? '#201f1e' : '#ffffff';
  }

  /**
   * Get filtered FAQ items based on selected category and search query
   */
  private _getFilteredItems(): IFaqItem[] {
    let filtered: IFaqItem[] = [];

    // First filter by category
    if (this._selectedCategory === 'All') {
      filtered = this._faqItems.slice();
    } else if (this._selectedCategory === this._topFaqCategoryKey) {
      let idx = 0;
      while (idx < this._faqItems.length) {
        if (this._faqItems[idx].isTopFaq) {
          filtered.push(this._faqItems[idx]);
        }
        idx++;
      }
    } else {
      let i = 0;
      while (i < this._faqItems.length) {
        if (this._faqItems[i].category === this._selectedCategory) {
          filtered.push(this._faqItems[i]);
        }
        i++;
      }
    }

    // Then filter by search query
    if (this._searchQuery && this._searchQuery.trim() !== '') {
      const query = this._searchQuery.toLowerCase().trim();
      const searchResults: IFaqItem[] = [];

      let j = 0;
      while (j < filtered.length) {
        const item = filtered[j];
        const titleMatch = item.title && item.title.toLowerCase().indexOf(query) !== -1;
        let answerMatch = false;

        if (this.properties.searchInAnswers && item.description) {
          // Strip HTML tags for search using DOM-based sanitization
          const plainDescription = this._stripHtmlTags(item.description);
          answerMatch = plainDescription.toLowerCase().indexOf(query) !== -1;
        }

        if (titleMatch || answerMatch) {
          searchResults.push(item);
        }
        j++;
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

    if (this._selectedCategory === this._topFaqCategoryKey) {
      let topCount = 0;
      let idx = 0;
      while (idx < this._faqItems.length) {
        if (this._faqItems[idx].isTopFaq) {
          topCount++;
        }
        idx++;
      }
      return topCount;
    }

    let count = 0;
    let i = 0;
    while (i < this._faqItems.length) {
      if (this._faqItems[i].category === this._selectedCategory) {
        count++;
      }
      i++;
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
      
      let i = 0;
      while (i < textNodes.length) {
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
        i++;
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
      const result = container.textContent || '';

      // Decode common HTML entities
      return this._decodeHtmlEntities(result);
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
   * Decode common HTML entities using split/join (ReDoS safe)
   */
  private _decodeHtmlEntities(text: string): string {
    // Use split/join instead of regex to avoid ReDoS
    // Process each entity directly to avoid for loop lint warning
    let result = text;
    result = result.split('&nbsp;').join(' ');
    result = result.split('&amp;').join('&');
    result = result.split('&lt;').join('<');
    result = result.split('&gt;').join('>');
    result = result.split('&quot;').join('"');
    result = result.split('&#39;').join("'");
    result = result.split('&#x27;').join("'");
    result = result.split('&apos;').join("'");
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

        // Build category tabs (All + optional Top FAQ + categories)
        const hasCategoryTabs = !!(this.properties.categoryColumn && this._categories.length > 0);
        const showTopFaqTab = this._isTopFaqEnabled();
        if (hasCategoryTabs || showTopFaqTab) {
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

          if (showTopFaqTab) {
            const isTopActive = this._selectedCategory === this._topFaqCategoryKey;
            const topActiveClass = isTopActive ? ' ' + styles.activeTab : '';
            const topFaqColor = this._getTopFaqColor(headerBgColor);
            const topBadgePalette = this._getTopFaqBadgePalette(topFaqColor);
            const badgeStyle = ' style="--topFaqBadgeBackground: ' + this._escapeHtml(topBadgePalette.badgeBackground) +
              '; --topFaqBadgeColor: ' + this._escapeHtml(topBadgePalette.badgeTextColor) +
              '; --topFaqBadgeBorder: ' + this._escapeHtml(topBadgePalette.badgeBorder) +
              '; --topFaqBadgeShadow: ' + this._escapeHtml(topBadgePalette.badgeShadow) + ';"';
            tabsArray.push(
              '<button class="' + styles.categoryTab + topActiveClass + '" ' +
              'data-category="' + this._topFaqCategoryKey + '" ' +
              'style="color: ' + (isTopActive ? headerBgColor : bodyText) + '; ' +
              'border-bottom-color: ' + (isTopActive ? headerBgColor : 'transparent') + ';">' +
              '<span class="' + styles.topFaqTab + '">' +
              '<span class="' + styles.topFaqBadge + '"' + badgeStyle + '>⭐</span>' +
              '<span class="' + styles.topFaqBadgeLabel + '">Top FAQ</span>' +
              '</span>' +
              '</button>'
            );
          }

          if (hasCategoryTabs) {
            let catIdx = 0;
            while (catIdx < this._categories.length) {
              const category = this._categories[catIdx];
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
              catIdx++;
            }
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
          let index = 0;
          while (index < filteredItems.length) {
            const item = filteredItems[index];
            let attachmentsHtml = '';

            if (item.attachments && item.attachments.length > 0) {
              const attachmentLinksArray: string[] = [];
              let attIdx = 0;
              while (attIdx < item.attachments.length) {
                const att = item.attachments[attIdx];
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
                attIdx++;
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
            index++;
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

      // Handle Escape key
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

    let tabIdx = 0;
    while (tabIdx < tabs.length) {
      const tab = tabs[tabIdx] as HTMLElement;
      (function(t: HTMLElement): void {
        t.addEventListener('click', function(): void {
          const category = t.dataset.category;
          if (category) {
            self._selectedCategory = category;
            self.render();
          }
        });
      })(tab);
      tabIdx++;
    }
  }

  private _attachEventListeners(): void {
    const faqItems = this.domElement.querySelectorAll('.' + styles.faqItem);
    const questions = this.domElement.querySelectorAll('.' + styles.faqQuestion);
    const self = this;

    let index = 0;
    while (index < questions.length) {
      const questionElement = questions[index] as HTMLElement;

      (function(qEl: HTMLElement, idx: number): void {
        qEl.addEventListener('mouseenter', function(): void {
          const hoverBg = qEl.dataset.hoverBg;
          if (hoverBg) {
            qEl.style.backgroundColor = hoverBg;
          }
        });

        qEl.addEventListener('mouseleave', function(): void {
          qEl.style.backgroundColor = '';
        });

        qEl.addEventListener('click', function(): void {
          const faqItem = faqItems[idx];

          if (!self.properties.allowMultipleExpanded) {
            let i = 0;
            while (i < faqItems.length) {
              if (i !== idx && faqItems[i].classList.contains(styles.expanded)) {
                faqItems[i].classList.remove(styles.expanded);
              }
              i++;
            }
          }

          if (faqItem.classList.contains(styles.expanded)) {
            faqItem.classList.remove(styles.expanded);
          } else {
            faqItem.classList.add(styles.expanded);
          }
        });
      })(questionElement, index);

      index++;
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
    return this._escapeHtml(description).split('\n').join('<br>');
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
      this.properties.categoryColumn || undefined,
      this.properties.topFaqColumn || undefined
    )
      .then((items: IFaqItem[]) => {
        this._faqItems = items;
        this._extractCategories();
        const hasTopFaqItems = this._isTopFaqEnabled() && this._hasTopFaqItems();
        if (hasTopFaqItems && this._selectedCategory === 'All') {
          this._selectedCategory = this._topFaqCategoryKey;
        }
        this._ensureValidCategorySelection();
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
      this.properties.topFaqColumn = '';
      this._faqItems = [];
      this._categories = [];
      this._selectedCategory = 'All';
      this._searchQuery = '';

      this._loadColumns().then(function(): void {
        self.context.propertyPane.refresh();
        self.render();
      });
    } else if ((propertyPath === 'titleColumn' || propertyPath === 'descriptionColumn' || propertyPath === 'categoryColumn' || propertyPath === 'topFaqColumn') && newValue !== oldValue) {
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

    let i = 0;
    while (i < this._lists.length) {
      options.push({ key: this._lists[i].id, text: this._lists[i].title });
      i++;
    }

    return options;
  }

  private _getTitleColumnOptions(): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = [
      { key: '', text: '-- Select a column --' }
    ];

    let i = 0;
    while (i < this._columns.length) {
      const col = this._columns[i];
      if (col.type === 'Text' || col.type === 'Note') {
        options.push({ key: col.internalName, text: col.title });
      }
      i++;
    }

    return options;
  }

  private _getDescriptionColumnOptions(): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = [
      { key: '', text: '-- Select a column --' }
    ];

    let i = 0;
    while (i < this._columns.length) {
      const col = this._columns[i];
      if (col.type === 'Text' || col.type === 'Note') {
        const typeLabel = col.type === 'Note' ? 'Multi-line' : 'Single-line';
        options.push({ key: col.internalName, text: col.title + ' (' + typeLabel + ')' });
      }
      i++;
    }

    return options;
  }

  private _getCategoryColumnOptions(): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = [
      { key: '', text: '-- No category (disable tabs) --' }
    ];

    let i = 0;
    while (i < this._columns.length) {
      const col = this._columns[i];
      if (col.type === 'Text' || col.type === 'Choice') {
        options.push({ key: col.internalName, text: col.title });
      }
      i++;
    }

    return options;
  }

  private _getTopFaqColumnOptions(): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = [
      { key: '', text: '-- No Top FAQ column --' }
    ];

    let i = 0;
    while (i < this._columns.length) {
      const col = this._columns[i];
      if (col.type === 'Boolean' || col.type === 'Choice' || col.type === 'Text') {
        options.push({ key: col.internalName, text: col.title });
      }
      i++;
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
                }),
                PropertyPaneTextField('topFaqColor', {
                  label: 'Top FAQ Accent Color',
                  placeholder: '#0078d4 or red',
                  description: 'Leave blank to inherit the theme color from the page.'
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
                }),
                PropertyPaneDropdown('topFaqColumn', {
                  label: 'Top FAQ Column (Yes/No)',
                  options: this._getTopFaqColumnOptions(),
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
