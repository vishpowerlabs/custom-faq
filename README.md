# CustomFAQ SPFx Web Part

A SharePoint Framework (SPFx) web part that displays SharePoint list items as an interactive FAQ accordion with **full theme support** including section background awareness.

![CustomFAQ Preview](./assets/preview.png)

## Features

### 🎨 Full Theme Support
- **SPFx Theme Tokens**: Uses `[theme: tokenName, default: #color]` syntax in SCSS for automatic theme color adaptation
- **Section Background Awareness**: Implements `ThemeProvider` to respond to section background color changes
- **CSS Variables Fallback**: Converts theme semantic colors to CSS custom properties for maximum compatibility
- **High Contrast Mode**: Full support for Windows High Contrast accessibility mode

### 🔧 Dynamic Configuration
- **List Selection**: Dropdown loads all available custom lists from the current site
- **Column Mapping**: Dynamically loads text-type columns after list selection
  - Title Column: Single-line or multi-line text
  - Description Column: Single-line or multi-line text (supports rich text)

### 📎 Attachment Support
- Automatically detects and displays list item attachments
- Clickable links open files in new tabs
- File icon indicators for easy recognition

### 🎯 Accordion Behavior
- Smooth expand/collapse animations
- Toggle option: Allow single or multiple items expanded
- Keyboard accessible (focus states, enter to expand)

## Prerequisites

- Node.js v16 or v18 (LTS versions)
- SharePoint Online or SharePoint 2019+
- SPFx 1.18.2 development environment

## Installation

1. **Clone/Download the project**

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Serve locally for development**
   ```bash
   gulp serve
   ```

4. **Build for production**
   ```bash
   gulp bundle --ship
   gulp package-solution --ship
   ```

5. **Deploy**
   - Upload `./sharepoint/solution/custom-faq.sppkg` to your App Catalog
   - Add the web part to a page

## Configuration

### Property Pane Options

| Property | Description |
|----------|-------------|
| **Webpart Title** | Header title displayed above the FAQ list |
| **Webpart Description** | Subtitle/description text in the header |
| **Select List** | Choose a SharePoint list as the data source |
| **Title Column** | Column to use for FAQ question/title |
| **Description Column** | Column to use for FAQ answer/description |
| **Allow Multiple Expanded** | Toggle to allow multiple items open simultaneously |

---

## SharePoint List Setup

### Creating the FAQ List

1. **Create a new Custom List** in your SharePoint site
2. **Add the following columns**:

| Column Name | Type | Required | Description |
|-------------|------|----------|-------------|
| Title | Single line of text | Yes | The FAQ question (default column) |
| Answer | Multiple lines of text | Yes | The FAQ answer (enable rich text for formatting) |
| Category | Choice | No | Optional category for organizing FAQs |
| SortOrder | Number | No | Optional numeric field for custom ordering |

3. **Enable Attachments** in List Settings → Advanced Settings → Attachments (if you want to include files)

### Importing Sample Data

A sample CSV file (`FAQ-List-Schema.csv`) is included with this project. To import:

1. Go to your SharePoint site
2. Create a new list or use an existing one
3. Use **Quick Edit** view or import via Power Automate/PnP PowerShell

---

## 🔐 List Permissions Setup

To ensure proper security while allowing all users to read FAQ items and contributors to manage only their own items, configure the following permissions:

### Step 1: Break Permission Inheritance

1. Navigate to your FAQ list
2. Go to **List Settings** → **Permissions for this list**
3. Click **Stop Inheriting Permissions** in the ribbon
4. Confirm the action

### Step 2: Configure Permission Levels

#### Everyone / Visitors (Read-Only Access)
All authenticated users should have **read-only** access to view FAQ items. They cannot create, edit, or delete any items.

1. Click **Grant Permissions** in the ribbon
2. Add the group: `Everyone` or `Everyone except external users`
3. Select permission level: **Read**
4. Click **Share**

> ⚠️ **Important**: Visitors with Read permission can only view items. They have no ability to create new items or modify existing ones.

#### Contributors (Edit Own Items Only)
Users who can submit or edit FAQ items should only modify their own entries.

1. Click **Grant Permissions**
2. Add the relevant users or groups (e.g., `FAQ Contributors`)
3. Select permission level: **Contribute**
4. Click **Share**

### Step 3: Enable Item-Level Permissions

This is the critical step that restricts contributors to only edit/delete their own items while visitors remain read-only:

1. Go to **List Settings**
2. Click **Advanced Settings**
3. Under **Item-level Permissions**, configure:

| Setting | Value | Description |
|---------|-------|-------------|
| **Read access** | Read all items | Everyone can read all FAQ items |
| **Create and Edit access** | Create items and edit items that were created by the user | Contributors can only modify their own items |

4. Click **OK** to save

> 📝 **Note**: These settings apply to users with Contribute permissions. Users with only Read permission are unaffected—they still cannot create or edit anything.

### Permission Matrix Summary

| User Role | Permission Level | Read All Items | Create Items | Edit Own Items | Edit All Items | Delete Own Items | Delete All Items |
|-----------|------------------|----------------|--------------|----------------|----------------|------------------|------------------|
| Visitors | Read | ✅ | ❌ | ❌ | ❌ | ❌ | ❌ |
| Contributors | Contribute | ✅ | ✅ | ✅ | ❌ | ✅ | ❌ |
| Owners/Admins | Full Control | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |

**Note:** Visitors have **Read-only** access. They cannot create, edit, or delete any items.



---

## Theme Integration

### CSS Theme Tokens

The web part uses SPFx theme tokens in SCSS:

```scss
.faqHeader {
  background-color: "[theme:themePrimary, default: #0078d4]";
}

.faqQuestionText {
  color: "[theme:bodyText, default: #323130]";
}
```

### ThemeProvider for Section Backgrounds

```typescript
// Consume ThemeProvider service
this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
this._themeVariant = this._themeProvider.tryGetTheme();

// Handle theme changes
this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);
```

### CSS Variables

Theme colors are also exposed as CSS custom properties:

```scss
.myElement {
  color: var(--bodyText, #323130);
  background: var(--neutralLighter, #f4f4f4);
}
```

## Available Theme Tokens

| Token | Usage |
|-------|-------|
| `themePrimary` | Primary brand color, links, buttons |
| `themeDark` | Darker variant for hover states |
| `bodyText` | Main text color |
| `bodySubtext` | Secondary text color |
| `bodyBackground` | Page/card background |
| `bodyDivider` | Borders and dividers |
| `link` / `linkHovered` | Link colors |
| `neutralLighter` | Light backgrounds |
| `neutralTertiary` | Tertiary elements |

## Project Structure

```
custom-faq/
├── config/
│   ├── config.json
│   ├── package-solution.json
│   └── serve.json
├── src/
│   └── webparts/
│       └── customFaq/
│           ├── CustomFaqWebPart.ts      # Main web part class
│           ├── CustomFaqWebPart.manifest.json
│           ├── CustomFaq.module.scss    # Themed styles
│           ├── services/
│           │   └── SharePointService.ts # SP REST API helpers
│           └── loc/
│               ├── en-us.js
│               └── mystrings.d.ts
├── FAQ-List-Schema.csv                  # Sample list data for import
├── package.json
├── tsconfig.json
└── gulpfile.js
```

## API Reference

### SharePointService Methods

```typescript
// Get all lists
getLists(): Promise<IListInfo[]>

// Get columns for a list
getListColumns(listId: string): Promise<IColumnInfo[]>

// Get items with attachments
getListItems(listId, titleColumn, descriptionColumn): Promise<IFaqItem[]>
```

## Browser Support

- Microsoft Edge (Chromium)
- Google Chrome
- Mozilla Firefox
- Safari

## Accessibility

- WCAG 2.1 AA compliant
- Keyboard navigation support
- Screen reader friendly
- High contrast mode support
- Focus indicators

## Troubleshooting

### Common Issues

| Issue | Solution |
|-------|----------|
| Lists not loading | Ensure user has at least Read permissions on the site |
| Columns not appearing | Only Text and Multi-line Text columns are shown |
| Attachments not displaying | Enable attachments in List Settings → Advanced Settings |
| Theme colors not applying | Ensure `supportsThemeVariants: true` is in manifest |
| Edit button not showing | User doesn't have Contribute permissions |

### Debug Mode

To enable debug logging, open browser console and run:
```javascript
localStorage.setItem('spfx-debug', 'true');
```

## License

MIT License

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## Support

For issues and feature requests, please create an issue in the repository.

---

## Changelog

### v1.0.0 (Initial Release)
- FAQ accordion display from SharePoint list
- Dynamic list and column selection
- Full SPFx theme support with section background awareness
- Attachment support with clickable links
- Single/multiple expand toggle option
- Responsive design
- Accessibility compliant
