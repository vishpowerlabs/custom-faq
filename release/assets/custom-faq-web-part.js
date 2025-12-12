define("a1b2c3d4-e5f6-7890-abcd-ef1234567890_0.0.1", ["@microsoft/sp-core-library","@microsoft/sp-property-pane","@microsoft/sp-webpart-base","@microsoft/sp-component-base","@microsoft/sp-http"], (__WEBPACK_EXTERNAL_MODULE__676__, __WEBPACK_EXTERNAL_MODULE__877__, __WEBPACK_EXTERNAL_MODULE__642__, __WEBPACK_EXTERNAL_MODULE__962__, __WEBPACK_EXTERNAL_MODULE__909__) => { return /******/ (() => { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ 858:
/*!*****************************************************!*\
  !*** ./lib/webparts/customFaq/CustomFaq.module.css ***!
  \*****************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var _node_modules_microsoft_sp_css_loader_node_modules_microsoft_load_themed_styles_lib_es6_index_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../../../node_modules/@microsoft/sp-css-loader/node_modules/@microsoft/load-themed-styles/lib-es6/index.js */ 323);
// Imports


_node_modules_microsoft_sp_css_loader_node_modules_microsoft_load_themed_styles_lib_es6_index_js__WEBPACK_IMPORTED_MODULE_0__.loadStyles(".customFaq_39a99b54{background-color:\"[theme:bodyBackground, default: #ffffff]\";background-color:var(--bodyBackground,#fff);border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,.08);font-family:Segoe UI,-apple-system,BlinkMacSystemFont,Roboto,Helvetica Neue,sans-serif;overflow:hidden}.faqHeader_39a99b54{background:\"[theme:themePrimary, default: #0078d4]\";background:var(--themePrimary,#0078d4);color:\"[theme:white, default: #ffffff]\";color:var(--white,#fff);padding:24px 28px}.faqTitle_39a99b54{font-weight:600;line-height:1.3;margin:0 0 4px}.faqDescription_39a99b54{line-height:1.4;margin:0;opacity:.9}.searchSection_39a99b54{background-color:\"[theme:neutralLighter, default: #f4f4f4]\";background-color:var(--neutralLighter,#f4f4f4);border-bottom:1px solid;border-bottom-color:var(--bodyDivider,#edebe9);padding:16px 20px}.searchContainer_39a99b54{max-width:100%;position:relative}.searchIcon_39a99b54{color:\"[theme:bodySubtext, default: #605e5c]\";color:var(--bodySubtext,#605e5c);height:16px;left:12px;pointer-events:none;position:absolute;top:50%;transform:translateY(-50%);width:16px}.searchInput_39a99b54{background-color:\"[theme:inputBackground, default: #ffffff]\";background-color:var(--inputBackground,#fff);border:1px solid;border-color:var(--inputBorder,#8a8886);border-radius:4px;color:\"[theme:bodyText, default: #323130]\";color:var(--bodyText,#323130);font-family:inherit;font-size:14px;padding:10px 40px;transition:border-color .2s,box-shadow .2s;width:100%}.searchInput:-ms-input-placeholder{color:\"[theme:inputPlaceholderText, default: #a19f9d]\";color:var(--inputPlaceholderText,#a19f9d)}.searchInput_39a99b54:-ms-input-placeholder{color:\"[theme:inputPlaceholderText, default: #a19f9d]\";color:var(--inputPlaceholderText,#a19f9d)}.searchInput_39a99b54::placeholder{color:\"[theme:inputPlaceholderText, default: #a19f9d]\";color:var(--inputPlaceholderText,#a19f9d)}.searchInput_39a99b54:focus{border-color:\"[theme:themePrimary, default: #0078d4]\";border-color:var(--themePrimary,#0078d4);box-shadow:0 0 0 1px var(--themePrimary,#0078d4);outline:0}.clearButton_39a99b54{align-items:center;background:0 0;border:none;border-radius:2px;color:\"[theme:bodySubtext, default: #605e5c]\";color:var(--bodySubtext,#605e5c);cursor:pointer;display:flex;justify-content:center;padding:4px;position:absolute;right:8px;top:50%;transform:translateY(-50%)}.clearButton_39a99b54:hover{background-color:\"[theme:neutralLight, default: #edebe9]\";background-color:var(--neutralLight,#edebe9)}.clearButton_39a99b54:focus{outline:2px solid;outline-color:var(--focusBorder,#0078d4);outline-offset:1px}.searchInHeader_39a99b54{margin-top:16px}.searchInHeader_39a99b54 .searchContainer_39a99b54{max-width:100%}.searchInHeader_39a99b54 .searchInput_39a99b54{background-color:hsla(0,0%,100%,.15);border-color:hsla(0,0%,100%,.3);color:#fff}.searchInHeader .searchInput:-ms-input-placeholder{color:hsla(0,0%,100%,.7)}.searchInHeader_39a99b54 .searchInput_39a99b54:-ms-input-placeholder{color:hsla(0,0%,100%,.7)}.searchInHeader_39a99b54 .searchInput_39a99b54::placeholder{color:hsla(0,0%,100%,.7)}.searchInHeader_39a99b54 .searchInput_39a99b54:focus{background-color:#fff;border-color:#fff;box-shadow:none;color:\"[theme:bodyText, default: #323130]\";color:var(--bodyText,#323130)}.searchInHeader .searchInput:focus:-ms-input-placeholder{color:\"[theme:inputPlaceholderText, default: #a19f9d]\";color:var(--inputPlaceholderText,#a19f9d)}.searchInHeader_39a99b54 .searchInput_39a99b54:focus:-ms-input-placeholder{color:\"[theme:inputPlaceholderText, default: #a19f9d]\";color:var(--inputPlaceholderText,#a19f9d)}.searchInHeader_39a99b54 .searchInput_39a99b54:focus::placeholder{color:\"[theme:inputPlaceholderText, default: #a19f9d]\";color:var(--inputPlaceholderText,#a19f9d)}.searchInHeader_39a99b54 .searchIcon_39a99b54{color:hsla(0,0%,100%,.7)}.searchInHeader_39a99b54 .searchContainer_39a99b54:focus-within .searchIcon_39a99b54,.searchInHeader_39a99b54 .searchInput_39a99b54:focus~.searchIcon_39a99b54{color:\"[theme:bodySubtext, default: #605e5c]\";color:var(--bodySubtext,#605e5c)}.searchInHeader_39a99b54 .clearButton_39a99b54{color:hsla(0,0%,100%,.7)}.searchInHeader_39a99b54 .clearButton_39a99b54:hover{background-color:hsla(0,0%,100%,.2)}.searchInHeader_39a99b54 .searchContainer_39a99b54:focus-within .clearButton_39a99b54,.searchInHeader_39a99b54 .searchInput_39a99b54:focus~.clearButton_39a99b54{color:\"[theme:bodySubtext, default: #605e5c]\";color:var(--bodySubtext,#605e5c)}.searchInHeader_39a99b54 .searchContainer_39a99b54:focus-within .clearButton_39a99b54:hover,.searchInHeader_39a99b54 .searchInput_39a99b54:focus~.clearButton_39a99b54:hover{background-color:\"[theme:neutralLight, default: #edebe9]\";background-color:var(--neutralLight,#edebe9)}.resultsInfo_39a99b54{background-color:\"[theme:neutralLighter, default: #f4f4f4]\";background-color:var(--neutralLighter,#f4f4f4);border-bottom:1px solid;border-bottom-color:var(--bodyDivider,#edebe9);color:\"[theme:bodySubtext, default: #605e5c]\";color:var(--bodySubtext,#605e5c);font-size:12px;padding:8px 20px}.resultsInfo_39a99b54 strong{color:\"[theme:bodyText, default: #323130]\";color:var(--bodyText,#323130)}.searchHighlight_39a99b54{background-color:#fff3cd;border-radius:2px;padding:1px 2px}.topFaqBadge_39a99b54{align-items:center;background-color:var(--topFaqBadgeBackground,#ffb900);border:1px solid var(--topFaqBadgeBorder,#c78600);border-radius:50%;box-shadow:0 6px 12px var(--topFaqBadgeShadow,rgba(0,0,0,.25));color:var(--topFaqBadgeColor,#fff);display:inline-flex;flex-shrink:0;font-size:16px;font-weight:700;height:28px;justify-content:center;width:28px}.topFaqTab_39a99b54{align-items:center;display:inline-flex;font-weight:600;gap:8px}.topFaqBadgeLabel_39a99b54{font-size:13px;letter-spacing:.2px}@media (forced-colors:active){.searchHighlight_39a99b54{background-color:Highlight;color:HighlightText}}.categoryTabs_39a99b54{-webkit-overflow-scrolling:touch;-ms-overflow-style:none;background-color:\"[theme:bodyBackground, default: #ffffff]\";background-color:var(--bodyBackground,#fff);border-bottom:1px solid;border-bottom-color:var(--neutralLight,#eaeaea);display:flex;flex-wrap:wrap;gap:4px;overflow-x:auto;padding:12px 20px;scrollbar-width:none}.categoryTabs_39a99b54::-webkit-scrollbar{display:none}.categoryTab_39a99b54{align-items:center;background:0 0;border:none;border-bottom:2px solid transparent;border-radius:4px 4px 0 0;color:\"[theme:bodyText, default: #323130]\";color:var(--bodyText,#323130);cursor:pointer;display:inline-flex;font-size:14px;font-weight:500;justify-content:center;padding:8px 16px;transition:all .2s ease;white-space:nowrap}.categoryTab_39a99b54:hover{background-color:\"[theme:neutralLighter, default: #f4f4f4]\";background-color:var(--neutralLighter,#f4f4f4)}.categoryTab_39a99b54:focus{outline:2px solid;outline-color:var(--focusBorder,#0078d4);outline-offset:-2px}.activeTab_39a99b54{background-color:\"[theme:neutralLighterAlt, default: #f8f8f8]\";background-color:var(--neutralLighterAlt,#f8f8f8);border-bottom-color:\"[theme:themePrimary, default: #0078d4]\";border-bottom-color:var(--themePrimary,#0078d4);color:\"[theme:themePrimary, default: #0078d4]\";color:var(--themePrimary,#0078d4);font-weight:600}.faqList_39a99b54{margin:0;padding:0}.faqItem_39a99b54{border-bottom:1px solid;border-bottom-color:var(--bodyDivider,#edebe9)}.faqItem_39a99b54:last-child{border-bottom:none}.faqQuestion_39a99b54{align-items:center;background-color:transparent;cursor:pointer;display:flex;justify-content:space-between;padding:20px 28px;transition:background-color .2s ease}.faqQuestion_39a99b54:hover{background-color:\"[theme:listItemBackgroundHovered, default: #f3f2f1]\";background-color:var(--listItemBackgroundHovered,#f3f2f1)}.faqQuestion_39a99b54:focus{outline:2px solid;outline-color:var(--focusBorder,#0078d4);outline-offset:-2px}.faqQuestionText_39a99b54{color:\"[theme:bodyText, default: #323130]\";color:var(--bodyText,#323130);flex:1;font-weight:600;padding-right:16px}.faqChevron_39a99b54{align-items:center;color:\"[theme:bodySubtext, default: #605e5c]\";color:var(--bodySubtext,#605e5c);display:flex;flex-shrink:0;height:24px;justify-content:center;transition:transform .3s ease;width:24px}.faqChevron_39a99b54 svg{height:16px;width:16px}.faqItem_39a99b54.expanded_39a99b54 .faqChevron_39a99b54{transform:rotate(180deg)}.faqAnswer_39a99b54{max-height:0;opacity:0;overflow:hidden;padding:0 28px;transition:max-height .3s ease,opacity .3s ease,padding .3s ease}.faqItem_39a99b54.expanded_39a99b54 .faqAnswer_39a99b54{max-height:2000px;opacity:1;padding:0 28px 20px}.faqAnswerContent_39a99b54{color:\"[theme:bodySubtext, default: #605e5c]\";color:var(--bodySubtext,#605e5c);line-height:1.6;margin-bottom:16px}.faqAnswerContent_39a99b54 p{margin:0 0 12px}.faqAnswerContent_39a99b54 p:last-child{margin-bottom:0}.faqAnswerContent_39a99b54 ol,.faqAnswerContent_39a99b54 ul{margin:0 0 12px;padding-left:24px}.faqAnswerContent_39a99b54 a{color:\"[theme:link, default: #0078d4]\";color:var(--link,#0078d4);text-decoration:none}.faqAnswerContent_39a99b54 a:hover{color:\"[theme:linkHovered, default: #004578]\";color:var(--linkHovered,#004578);text-decoration:underline}.attachments_39a99b54{background-color:\"[theme:neutralLighter, default: #f4f4f4]\";background-color:var(--neutralLighter,#f4f4f4);border-radius:6px;margin-top:12px;padding:12px 16px}.attachmentsLabel_39a99b54{align-items:center;color:\"[theme:neutralTertiary, default: #a6a6a6]\";color:var(--neutralTertiary,#a6a6a6);display:flex;font-size:.75rem;font-weight:600;gap:6px;letter-spacing:.5px;margin-bottom:8px;text-transform:uppercase}.attachmentsLabel_39a99b54 svg{flex-shrink:0}.attachmentLink_39a99b54{align-items:center;color:\"[theme:link, default: #0078d4]\";color:var(--link,#0078d4);display:flex;font-size:.875rem;gap:8px;padding:6px 0;text-decoration:none;transition:color .2s ease}.attachmentLink_39a99b54:hover{color:\"[theme:linkHovered, default: #004578]\";color:var(--linkHovered,#004578);text-decoration:underline}.attachmentLink_39a99b54:focus{outline:2px solid;outline-color:var(--focusBorder,#0078d4);outline-offset:2px}.attachmentIcon_39a99b54{align-items:center;display:flex;flex-shrink:0;height:20px;justify-content:center;width:20px}.attachmentIcon_39a99b54 svg{height:16px;width:16px}.emptyState_39a99b54{align-items:center;color:\"[theme:bodySubtext, default: #605e5c]\";color:var(--bodySubtext,#605e5c);display:flex;flex-direction:column;justify-content:center;padding:48px 24px;text-align:center}.emptyState_39a99b54 svg{margin-bottom:16px;opacity:.6}.emptyState_39a99b54 p{font-size:.9375rem;line-height:1.5;margin:0;max-width:300px}.loading_39a99b54{align-items:center;display:flex;justify-content:center;padding:48px 24px}.loading_39a99b54:after{animation:spin_39a99b54 1s linear infinite;border:3px solid;border-radius:50%;border-top:3px solid \"[theme:themeprimary, default: #0078d4]\";content:\"\";height:32px;width:32px}@keyframes spin_39a99b54{to{transform:rotate(1turn)}}@media (forced-colors:active){.faqQuestion_39a99b54{border:1px solid ButtonText}.faqItem_39a99b54.expanded_39a99b54 .faqQuestion_39a99b54,.faqQuestion_39a99b54:hover{background-color:Highlight;color:HighlightText}.attachmentLink_39a99b54,.attachmentLink_39a99b54:hover{color:LinkText}.categoryTab_39a99b54{border:1px solid ButtonText}.activeTab_39a99b54,.categoryTab_39a99b54:hover{background-color:Highlight;color:HighlightText}.searchInput_39a99b54{border:1px solid ButtonText}}.faqQuestion_39a99b54:focus-visible{outline:2px solid;outline-color:var(--focusBorder,#0078d4);outline-offset:-2px}.attachmentLink_39a99b54:focus-visible{outline:2px solid;outline-color:var(--focusBorder,#0078d4);outline-offset:2px}.categoryTab_39a99b54:focus-visible{outline:2px solid;outline-color:var(--focusBorder,#0078d4);outline-offset:-2px}@media screen and (max-width:640px){.faqHeader_39a99b54{padding:20px}.faqQuestion_39a99b54{padding:16px 20px}.faqItem_39a99b54.expanded_39a99b54 .faqAnswer_39a99b54{padding:0 20px 16px}.attachments_39a99b54{padding:10px 12px}.emptyState_39a99b54{padding:32px 20px}.categoryTabs_39a99b54{padding:10px 16px}.categoryTab_39a99b54{font-size:13px;padding:6px 12px}.searchSection_39a99b54{padding:12px 16px}.searchInHeader_39a99b54{margin-top:12px}.resultsInfo_39a99b54{padding:6px 16px}}@media print{.customFaq_39a99b54{border:1px solid #ccc;box-shadow:none}.faqHeader_39a99b54{background:#333!important;-webkit-print-color-adjust:exact;print-color-adjust:exact}.faqAnswer_39a99b54{max-height:none!important;opacity:1!important;padding:0 28px 20px!important}.faqChevron_39a99b54{display:none}.faqQuestion_39a99b54{cursor:default}.categoryTabs_39a99b54,.resultsInfo_39a99b54,.searchInHeader_39a99b54,.searchSection_39a99b54{display:none}.searchHighlight_39a99b54{background-color:#ff0!important;-webkit-print-color-adjust:exact;print-color-adjust:exact}}\n/*# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJzb3VyY2VzIjpbImZpbGU6Ly8vVXNlcnMvdmlzaG51L0Rlc2t0b3AvY29kZWJhc2UvY3VzdG9tLWZhcS9zcmMvd2VicGFydHMvY3VzdG9tRmFxL0N1c3RvbUZhcS5tb2R1bGUuc2NzcyJdLCJuYW1lcyI6W10sIm1hcHBpbmdzIjoiQUFPQSxvQkFNRSwyREFBQSxDQUNBLDJDQUFBLENBTEEsaUJBQUEsQ0FFQSxvQ0FBQSxDQUhBLHNGQUFBLENBRUEsZUFJQSxDQU9GLG9CQUtFLG1EQUFBLENBQ0Esc0NBQUEsQ0FKQSx1Q0FBQSxDQUNBLHVCQUFBLENBRkEsaUJBS0EsQ0FHRixtQkFDRSxlQUFBLENBRUEsZUFBQSxDQURBLGNBQ0EsQ0FHRix5QkFHRSxlQUFBLENBRkEsUUFBQSxDQUNBLFVBQ0EsQ0FPRix3QkFLRSwyREFBQSxDQUNBLDhDQUFBLENBSEEsdUJBQUEsQ0FBQSw4Q0FBQSxDQUZBLGlCQUtBLENBR0YsMEJBRUUsY0FBQSxDQURBLGlCQUNBLENBR0YscUJBU0UsNkNBQUEsQ0FDQSxnQ0FBQSxDQUpBLFdBQUEsQ0FKQSxTQUFBLENBS0EsbUJBQUEsQ0FOQSxpQkFBQSxDQUVBLE9BQUEsQ0FDQSwwQkFBQSxDQUNBLFVBS0EsQ0FHRixzQkFVRSw0REFBQSxDQUNBLDRDQUFBLENBUkEsZ0JBQUEsQ0FDQSx1Q0FBQSxDQUNBLGlCQUFBLENBUUEsMENBQUEsQ0FDQSw2QkFBQSxDQVBBLG1CQUFBLENBREEsY0FBQSxDQUpBLGlCQUFBLENBTUEsMENBQUEsQ0FQQSxVQWFBLENBRUEsbUNBQ0Usc0RBQUEsQ0FDQSx5Q0FBQSxDQUZGLDRDQUNFLHNEQUFBLENBQ0EseUNBQUEsQ0FGRixtQ0FDRSxzREFBQSxDQUNBLHlDQUFBLENBR0YsNEJBRUUscURBQUEsQ0FDQSx3Q0FBQSxDQUNBLGdEQUFBLENBSEEsU0FHQSxDQUlKLHNCQVdFLGtCQUFBLENBTkEsY0FBQSxDQUNBLFdBQUEsQ0FFQSxpQkFBQSxDQU1BLDZDQUFBLENBQ0EsZ0NBQUEsQ0FOQSxjQUFBLENBQ0EsWUFBQSxDQUVB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 */", true);

// Exports
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = ({
  customFaq_39a99b54: "customFaq_39a99b54",
  faqHeader_39a99b54: "faqHeader_39a99b54",
  faqTitle_39a99b54: "faqTitle_39a99b54",
  faqDescription_39a99b54: "faqDescription_39a99b54",
  searchSection_39a99b54: "searchSection_39a99b54",
  searchContainer_39a99b54: "searchContainer_39a99b54",
  searchIcon_39a99b54: "searchIcon_39a99b54",
  searchInput_39a99b54: "searchInput_39a99b54",
  searchInput: "searchInput",
  clearButton_39a99b54: "clearButton_39a99b54",
  searchInHeader_39a99b54: "searchInHeader_39a99b54",
  searchInHeader: "searchInHeader",
  resultsInfo_39a99b54: "resultsInfo_39a99b54",
  searchHighlight_39a99b54: "searchHighlight_39a99b54",
  topFaqBadge_39a99b54: "topFaqBadge_39a99b54",
  topFaqTab_39a99b54: "topFaqTab_39a99b54",
  topFaqBadgeLabel_39a99b54: "topFaqBadgeLabel_39a99b54",
  categoryTabs_39a99b54: "categoryTabs_39a99b54",
  categoryTab_39a99b54: "categoryTab_39a99b54",
  activeTab_39a99b54: "activeTab_39a99b54",
  faqList_39a99b54: "faqList_39a99b54",
  faqItem_39a99b54: "faqItem_39a99b54",
  faqQuestion_39a99b54: "faqQuestion_39a99b54",
  faqQuestionText_39a99b54: "faqQuestionText_39a99b54",
  faqChevron_39a99b54: "faqChevron_39a99b54",
  expanded_39a99b54: "expanded_39a99b54",
  faqAnswer_39a99b54: "faqAnswer_39a99b54",
  faqAnswerContent_39a99b54: "faqAnswerContent_39a99b54",
  attachments_39a99b54: "attachments_39a99b54",
  attachmentsLabel_39a99b54: "attachmentsLabel_39a99b54",
  attachmentLink_39a99b54: "attachmentLink_39a99b54",
  attachmentIcon_39a99b54: "attachmentIcon_39a99b54",
  emptyState_39a99b54: "emptyState_39a99b54",
  loading_39a99b54: "loading_39a99b54",
  spin_39a99b54: "spin_39a99b54"
});


/***/ }),

/***/ 426:
/*!*********************************************************!*\
  !*** ./lib/webparts/customFaq/CustomFaq.module.scss.js ***!
  \*********************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
__webpack_require__(/*! ./CustomFaq.module.css */ 858);
var styles = {
    customFaq: 'customFaq_39a99b54',
    faqHeader: 'faqHeader_39a99b54',
    faqTitle: 'faqTitle_39a99b54',
    faqDescription: 'faqDescription_39a99b54',
    searchSection: 'searchSection_39a99b54',
    searchContainer: 'searchContainer_39a99b54',
    searchIcon: 'searchIcon_39a99b54',
    searchInput: 'searchInput_39a99b54',
    clearButton: 'clearButton_39a99b54',
    searchInHeader: 'searchInHeader_39a99b54',
    resultsInfo: 'resultsInfo_39a99b54',
    searchHighlight: 'searchHighlight_39a99b54',
    topFaqBadge: 'topFaqBadge_39a99b54',
    topFaqTab: 'topFaqTab_39a99b54',
    topFaqBadgeLabel: 'topFaqBadgeLabel_39a99b54',
    categoryTabs: 'categoryTabs_39a99b54',
    categoryTab: 'categoryTab_39a99b54',
    activeTab: 'activeTab_39a99b54',
    faqList: 'faqList_39a99b54',
    faqItem: 'faqItem_39a99b54',
    faqQuestion: 'faqQuestion_39a99b54',
    faqQuestionText: 'faqQuestionText_39a99b54',
    faqChevron: 'faqChevron_39a99b54',
    expanded: 'expanded_39a99b54',
    faqAnswer: 'faqAnswer_39a99b54',
    faqAnswerContent: 'faqAnswerContent_39a99b54',
    attachments: 'attachments_39a99b54',
    attachmentsLabel: 'attachmentsLabel_39a99b54',
    attachmentLink: 'attachmentLink_39a99b54',
    attachmentIcon: 'attachmentIcon_39a99b54',
    emptyState: 'emptyState_39a99b54',
    loading: 'loading_39a99b54',
    spin: 'spin_39a99b54'
};
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (styles);


/***/ }),

/***/ 666:
/*!**************************************************************!*\
  !*** ./lib/webparts/customFaq/services/SharePointService.js ***!
  \**************************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   SharePointService: () => (/* binding */ SharePointService)
/* harmony export */ });
/* harmony import */ var _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @microsoft/sp-http */ 909);
/* harmony import */ var _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__);

/**
 * SharePoint Service - Helper methods for retrieving lists, columns, and items
 */
var SharePointService = /** @class */ (function () {
    function SharePointService(context) {
        this._context = context;
        this._siteUrl = context.pageContext.web.absoluteUrl;
    }
    /**
     * Get all lists from the current site
     * Filters out hidden lists and system lists
     */
    SharePointService.prototype.getLists = function () {
        var endpoint = this._siteUrl + '/_api/web/lists?$filter=Hidden eq false and BaseTemplate eq 100&$select=Id,Title&$orderby=Title';
        return this._context.spHttpClient.get(endpoint, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__.SPHttpClient.configurations.v1)
            .then(function (response) {
            if (!response.ok) {
                throw new Error('Error fetching lists: ' + response.statusText);
            }
            return response.json();
        })
            .then(function (data) {
            var lists = [];
            var i = 0;
            while (i < data.value.length) {
                lists.push({
                    id: data.value[i].Id,
                    title: data.value[i].Title
                });
                i++;
            }
            return lists;
        });
    };
    /**
     * Get columns for a specific list
     * Returns Text, Note, Choice, and Boolean columns
     */
    SharePointService.prototype.getListColumns = function (listId) {
        // Filter for Text, Note, Choice, and Boolean field types
        var endpoint = this._siteUrl + '/_api/web/lists(guid\'' + listId + '\')/fields?$filter=(TypeAsString eq \'Text\' or TypeAsString eq \'Note\' or TypeAsString eq \'Choice\' or TypeAsString eq \'Boolean\') and Hidden eq false and ReadOnlyField eq false&$select=Id,InternalName,Title,TypeAsString&$orderby=Title';
        return this._context.spHttpClient.get(endpoint, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__.SPHttpClient.configurations.v1)
            .then(function (response) {
            if (!response.ok) {
                throw new Error('Error fetching columns: ' + response.statusText);
            }
            return response.json();
        })
            .then(function (data) {
            var columns = [];
            var i = 0;
            while (i < data.value.length) {
                columns.push({
                    id: data.value[i].Id,
                    internalName: data.value[i].InternalName,
                    title: data.value[i].Title,
                    type: data.value[i].TypeAsString
                });
                i++;
            }
            return columns;
        });
    };
    /**
     * Get all items from a list with specified columns
     * Also fetches attachments for each item
     */
    SharePointService.prototype.getListItems = function (listId, titleColumn, descriptionColumn, categoryColumn, topFaqColumn) {
        // Build select query - always include Id and Attachments
        var selectFields = ['Id', titleColumn];
        if (descriptionColumn !== titleColumn) {
            selectFields.push(descriptionColumn);
        }
        if (categoryColumn && categoryColumn !== titleColumn && categoryColumn !== descriptionColumn) {
            selectFields.push(categoryColumn);
        }
        if (topFaqColumn && selectFields.indexOf(topFaqColumn) === -1) {
            selectFields.push(topFaqColumn);
        }
        selectFields.push('Attachments');
        var endpoint = this._siteUrl + '/_api/web/lists(guid\'' + listId + '\')/items?$select=' + selectFields.join(',') + '&$top=500';
        var self = this;
        return this._context.spHttpClient.get(endpoint, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__.SPHttpClient.configurations.v1)
            .then(function (response) {
            if (!response.ok) {
                throw new Error('Error fetching items: ' + response.statusText);
            }
            return response.json();
        })
            .then(function (data) {
            // Process items and fetch attachments
            var itemPromises = [];
            var i = 0;
            while (i < data.value.length) {
                var item = data.value[i];
                itemPromises.push(self._processItem(item, listId, titleColumn, descriptionColumn, categoryColumn, topFaqColumn));
                i++;
            }
            return Promise.all(itemPromises);
        })
            .then(function (items) {
            // Filter out items without titles
            var filteredItems = [];
            var i = 0;
            while (i < items.length) {
                if (items[i].title && items[i].title.trim() !== '') {
                    filteredItems.push(items[i]);
                }
                i++;
            }
            return filteredItems;
        });
    };
    /**
     * Process a single item and fetch its attachments
     */
    SharePointService.prototype._processItem = function (item, listId, titleColumn, descriptionColumn, categoryColumn, topFaqColumn) {
        var self = this;
        // Determine if item is marked as Top FAQ
        var isTopFaq = false;
        if (topFaqColumn && item[topFaqColumn] !== undefined) {
            var topFaqValue = item[topFaqColumn];
            // Handle Boolean (Yes/No) columns
            if (typeof topFaqValue === 'boolean') {
                isTopFaq = topFaqValue === true;
            }
            // Handle Choice columns with "Yes"/"No" values
            else if (typeof topFaqValue === 'string') {
                var lowerValue = topFaqValue.trim().toLowerCase();
                isTopFaq = lowerValue === 'yes' || lowerValue === 'true' || lowerValue === '1' || lowerValue === 'y';
            }
            // Handle numeric values (1 = yes)
            else if (typeof topFaqValue === 'number') {
                isTopFaq = topFaqValue === 1;
            }
        }
        var faqItem = {
            id: item.Id,
            title: item[titleColumn] || '',
            description: item[descriptionColumn] || '',
            category: categoryColumn ? (item[categoryColumn] || '') : '',
            isTopFaq: isTopFaq,
            attachments: []
        };
        // If item has attachments, fetch them
        if (item.Attachments) {
            return self._getItemAttachments(listId, item.Id)
                .then(function (attachments) {
                faqItem.attachments = attachments;
                return faqItem;
            });
        }
        return Promise.resolve(faqItem);
    };
    /**
     * Get attachments for a specific list item
     */
    SharePointService.prototype._getItemAttachments = function (listId, itemId) {
        var endpoint = this._siteUrl + '/_api/web/lists(guid\'' + listId + '\')/items(' + itemId + ')/AttachmentFiles';
        var self = this;
        return this._context.spHttpClient.get(endpoint, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__.SPHttpClient.configurations.v1)
            .then(function (response) {
            if (!response.ok) {
                console.warn('Could not fetch attachments for item ' + itemId);
                return { value: [] };
            }
            return response.json();
        })
            .then(function (data) {
            var attachments = [];
            var baseUrl = self._context.pageContext.web.absoluteUrl.split('/').slice(0, 3).join('/');
            var i = 0;
            while (i < data.value.length) {
                attachments.push({
                    fileName: data.value[i].FileName,
                    url: baseUrl + data.value[i].ServerRelativeUrl
                });
                i++;
            }
            return attachments;
        })
            .catch(function (error) {
            console.warn('Error fetching attachments:', error);
            return [];
        });
    };
    /**
     * Get a specific list by ID
     */
    SharePointService.prototype.getListById = function (listId) {
        var endpoint = this._siteUrl + '/_api/web/lists(guid\'' + listId + '\')?$select=Id,Title';
        return this._context.spHttpClient.get(endpoint, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__.SPHttpClient.configurations.v1)
            .then(function (response) {
            if (!response.ok) {
                return null;
            }
            return response.json();
        })
            .then(function (data) {
            if (!data) {
                return null;
            }
            return {
                id: data.Id,
                title: data.Title
            };
        })
            .catch(function () {
            return null;
        });
    };
    /**
     * Check if a column exists in a list
     */
    SharePointService.prototype.columnExists = function (listId, columnInternalName) {
        var endpoint = this._siteUrl + '/_api/web/lists(guid\'' + listId + '\')/fields?$filter=InternalName eq \'' + columnInternalName + '\'&$select=Id';
        return this._context.spHttpClient.get(endpoint, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__.SPHttpClient.configurations.v1)
            .then(function (response) {
            if (!response.ok) {
                return false;
            }
            return response.json();
        })
            .then(function (data) {
            return data.value && data.value.length > 0;
        })
            .catch(function () {
            return false;
        });
    };
    return SharePointService;
}());



/***/ }),

/***/ 323:
/*!***********************************************************************************************************!*\
  !*** ./node_modules/@microsoft/sp-css-loader/node_modules/@microsoft/load-themed-styles/lib-es6/index.js ***!
  \***********************************************************************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   ClearStyleOptions: () => (/* binding */ ClearStyleOptions),
/* harmony export */   Mode: () => (/* binding */ Mode),
/* harmony export */   clearStyles: () => (/* binding */ clearStyles),
/* harmony export */   configureLoadStyles: () => (/* binding */ configureLoadStyles),
/* harmony export */   configureRunMode: () => (/* binding */ configureRunMode),
/* harmony export */   detokenize: () => (/* binding */ detokenize),
/* harmony export */   flush: () => (/* binding */ flush),
/* harmony export */   loadStyles: () => (/* binding */ loadStyles),
/* harmony export */   loadTheme: () => (/* binding */ loadTheme),
/* harmony export */   splitStyles: () => (/* binding */ splitStyles)
/* harmony export */ });
// Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
// See LICENSE in the project root for license information.
var __assign = (undefined && undefined.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
/**
 * In sync mode, styles are registered as style elements synchronously with loadStyles() call.
 * In async mode, styles are buffered and registered as batch in async timer for performance purpose.
 */
var Mode;
(function (Mode) {
    Mode[Mode["sync"] = 0] = "sync";
    Mode[Mode["async"] = 1] = "async";
})(Mode || (Mode = {}));
/**
 * Themable styles and non-themable styles are tracked separately
 * Specify ClearStyleOptions when calling clearStyles API to specify which group of registered styles should be cleared.
 */
var ClearStyleOptions;
(function (ClearStyleOptions) {
    /** only themable styles will be cleared */
    ClearStyleOptions[ClearStyleOptions["onlyThemable"] = 1] = "onlyThemable";
    /** only non-themable styles will be cleared */
    ClearStyleOptions[ClearStyleOptions["onlyNonThemable"] = 2] = "onlyNonThemable";
    /** both themable and non-themable styles will be cleared */
    ClearStyleOptions[ClearStyleOptions["all"] = 3] = "all";
})(ClearStyleOptions || (ClearStyleOptions = {}));
// Store the theming state in __themeState__ global scope for reuse in the case of duplicate
// load-themed-styles hosted on the page.
var _root = typeof window === 'undefined' ? __webpack_require__.g : window; // eslint-disable-line @typescript-eslint/no-explicit-any
// Nonce string to inject into script tag if one provided. This is used in CSP (Content Security Policy).
var _styleNonce = _root && _root.CSPSettings && _root.CSPSettings.nonce;
var _themeState = initializeThemeState();
/**
 * Matches theming tokens. For example, "[theme: themeSlotName, default: #FFF]" (including the quotes).
 */
var _themeTokenRegex = /[\'\"]\[theme:\s*([A-Za-z0-9_-]+)\s*(?:,\s*default:\s*([^\]\r\n]+?))?\s*\][\'\"]/g;
var now = function () {
    return typeof performance !== 'undefined' && !!performance.now ? performance.now() : Date.now();
};
function measure(func) {
    var start = now();
    func();
    var end = now();
    _themeState.perf.duration += end - start;
}
/**
 * initialize global state object
 */
function initializeThemeState() {
    var state = _root.__themeState__ || {
        theme: undefined,
        lastStyleElement: undefined,
        registeredStyles: []
    };
    if (!state.runState) {
        state = __assign(__assign({}, state), { perf: {
                count: 0,
                duration: 0
            }, runState: {
                flushTimer: 0,
                mode: Mode.sync,
                buffer: []
            } });
    }
    if (!state.registeredThemableStyles) {
        state = __assign(__assign({}, state), { registeredThemableStyles: [] });
    }
    _root.__themeState__ = state;
    return state;
}
/**
 * Loads a set of style text. If it is registered too early, we will register it when the window.load
 * event is fired.
 * @param {string | ThemableArray} styles Themable style text to register.
 * @param {boolean} loadAsync When true, always load styles in async mode, irrespective of current sync mode.
 */
function loadStyles(styles, loadAsync) {
    if (loadAsync === void 0) { loadAsync = false; }
    measure(function () {
        var styleParts = Array.isArray(styles) ? styles : splitStyles(styles);
        var _a = _themeState.runState, mode = _a.mode, buffer = _a.buffer, flushTimer = _a.flushTimer;
        if (loadAsync || mode === Mode.async) {
            buffer.push(styleParts);
            if (!flushTimer) {
                _themeState.runState.flushTimer = asyncLoadStyles();
            }
        }
        else {
            applyThemableStyles(styleParts);
        }
    });
}
/**
 * Allows for customizable loadStyles logic. e.g. for server side rendering application
 * @param {(processedStyles: string, rawStyles?: string | ThemableArray) => void}
 * a loadStyles callback that gets called when styles are loaded or reloaded
 */
function configureLoadStyles(loadStylesFn) {
    _themeState.loadStyles = loadStylesFn;
}
/**
 * Configure run mode of load-themable-styles
 * @param mode load-themable-styles run mode, async or sync
 */
function configureRunMode(mode) {
    _themeState.runState.mode = mode;
}
/**
 * external code can call flush to synchronously force processing of currently buffered styles
 */
function flush() {
    measure(function () {
        var styleArrays = _themeState.runState.buffer.slice();
        _themeState.runState.buffer = [];
        var mergedStyleArray = [].concat.apply([], styleArrays);
        if (mergedStyleArray.length > 0) {
            applyThemableStyles(mergedStyleArray);
        }
    });
}
/**
 * register async loadStyles
 */
function asyncLoadStyles() {
    // Use "self" to distinguish conflicting global typings for setTimeout() from lib.dom.d.ts vs Jest's @types/node
    // https://github.com/jestjs/jest/issues/14418
    return self.setTimeout(function () {
        _themeState.runState.flushTimer = 0;
        flush();
    }, 0);
}
/**
 * Loads a set of style text. If it is registered too early, we will register it when the window.load event
 * is fired.
 * @param {string} styleText Style to register.
 * @param {IStyleRecord} styleRecord Existing style record to re-apply.
 */
function applyThemableStyles(stylesArray, styleRecord) {
    if (_themeState.loadStyles) {
        _themeState.loadStyles(resolveThemableArray(stylesArray).styleString, stylesArray);
    }
    else {
        registerStyles(stylesArray);
    }
}
/**
 * Registers a set theme tokens to find and replace. If styles were already registered, they will be
 * replaced.
 * @param {theme} theme JSON object of theme tokens to values.
 */
function loadTheme(theme) {
    _themeState.theme = theme;
    // reload styles.
    reloadStyles();
}
/**
 * Clear already registered style elements and style records in theme_State object
 * @param option - specify which group of registered styles should be cleared.
 * Default to be both themable and non-themable styles will be cleared
 */
function clearStyles(option) {
    if (option === void 0) { option = ClearStyleOptions.all; }
    if (option === ClearStyleOptions.all || option === ClearStyleOptions.onlyNonThemable) {
        clearStylesInternal(_themeState.registeredStyles);
        _themeState.registeredStyles = [];
    }
    if (option === ClearStyleOptions.all || option === ClearStyleOptions.onlyThemable) {
        clearStylesInternal(_themeState.registeredThemableStyles);
        _themeState.registeredThemableStyles = [];
    }
}
function clearStylesInternal(records) {
    records.forEach(function (styleRecord) {
        var styleElement = styleRecord && styleRecord.styleElement;
        if (styleElement && styleElement.parentElement) {
            styleElement.parentElement.removeChild(styleElement);
        }
    });
}
/**
 * Reloads styles.
 */
function reloadStyles() {
    if (_themeState.theme) {
        var themableStyles = [];
        for (var _i = 0, _a = _themeState.registeredThemableStyles; _i < _a.length; _i++) {
            var styleRecord = _a[_i];
            themableStyles.push(styleRecord.themableStyle);
        }
        if (themableStyles.length > 0) {
            clearStyles(ClearStyleOptions.onlyThemable);
            applyThemableStyles([].concat.apply([], themableStyles));
        }
    }
}
/**
 * Find theme tokens and replaces them with provided theme values.
 * @param {string} styles Tokenized styles to fix.
 */
function detokenize(styles) {
    if (styles) {
        styles = resolveThemableArray(splitStyles(styles)).styleString;
    }
    return styles;
}
/**
 * Resolves ThemingInstruction objects in an array and joins the result into a string.
 * @param {ThemableArray} splitStyleArray ThemableArray to resolve and join.
 */
function resolveThemableArray(splitStyleArray) {
    var theme = _themeState.theme;
    var themable = false;
    // Resolve the array of theming instructions to an array of strings.
    // Then join the array to produce the final CSS string.
    var resolvedArray = (splitStyleArray || []).map(function (currentValue) {
        var themeSlot = currentValue.theme;
        if (themeSlot) {
            themable = true;
            // A theming annotation. Resolve it.
            var themedValue = theme ? theme[themeSlot] : undefined;
            var defaultValue = currentValue.defaultValue || 'inherit';
            // Warn to console if we hit an unthemed value even when themes are provided, but only if "DEBUG" is true.
            // Allow the themedValue to be undefined to explicitly request the default value.
            if (theme &&
                !themedValue &&
                console &&
                !(themeSlot in theme) &&
                "boolean" !== 'undefined' &&
                true) {
                // eslint-disable-next-line no-console
                console.warn("Theming value not provided for \"".concat(themeSlot, "\". Falling back to \"").concat(defaultValue, "\"."));
            }
            return themedValue || defaultValue;
        }
        else {
            // A non-themable string. Preserve it.
            return currentValue.rawString;
        }
    });
    return {
        styleString: resolvedArray.join(''),
        themable: themable
    };
}
/**
 * Split tokenized CSS into an array of strings and theme specification objects
 * @param {string} styles Tokenized styles to split.
 */
function splitStyles(styles) {
    var result = [];
    if (styles) {
        var pos = 0; // Current position in styles.
        var tokenMatch = void 0;
        while ((tokenMatch = _themeTokenRegex.exec(styles))) {
            var matchIndex = tokenMatch.index;
            if (matchIndex > pos) {
                result.push({
                    rawString: styles.substring(pos, matchIndex)
                });
            }
            result.push({
                theme: tokenMatch[1],
                defaultValue: tokenMatch[2] // May be undefined
            });
            // index of the first character after the current match
            pos = _themeTokenRegex.lastIndex;
        }
        // Push the rest of the string after the last match.
        result.push({
            rawString: styles.substring(pos)
        });
    }
    return result;
}
/**
 * Registers a set of style text. If it is registered too early, we will register it when the
 * window.load event is fired.
 * @param {ThemableArray} styleArray Array of IThemingInstruction objects to register.
 * @param {IStyleRecord} styleRecord May specify a style Element to update.
 */
function registerStyles(styleArray) {
    if (typeof document === 'undefined') {
        return;
    }
    var head = document.getElementsByTagName('head')[0];
    var styleElement = document.createElement('style');
    var _a = resolveThemableArray(styleArray), styleString = _a.styleString, themable = _a.themable;
    styleElement.setAttribute('data-load-themed-styles', 'true');
    if (_styleNonce) {
        styleElement.setAttribute('nonce', _styleNonce);
    }
    styleElement.appendChild(document.createTextNode(styleString));
    _themeState.perf.count++;
    head.appendChild(styleElement);
    var ev = document.createEvent('HTMLEvents');
    ev.initEvent('styleinsert', true /* bubbleEvent */, false /* cancelable */);
    ev.args = {
        newStyle: styleElement
    };
    document.dispatchEvent(ev);
    var record = {
        styleElement: styleElement,
        themableStyle: styleArray
    };
    if (themable) {
        _themeState.registeredThemableStyles.push(record);
    }
    else {
        _themeState.registeredStyles.push(record);
    }
}


/***/ }),

/***/ 962:
/*!***********************************************!*\
  !*** external "@microsoft/sp-component-base" ***!
  \***********************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__962__;

/***/ }),

/***/ 676:
/*!*********************************************!*\
  !*** external "@microsoft/sp-core-library" ***!
  \*********************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__676__;

/***/ }),

/***/ 909:
/*!*************************************!*\
  !*** external "@microsoft/sp-http" ***!
  \*************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__909__;

/***/ }),

/***/ 877:
/*!**********************************************!*\
  !*** external "@microsoft/sp-property-pane" ***!
  \**********************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__877__;

/***/ }),

/***/ 642:
/*!*********************************************!*\
  !*** external "@microsoft/sp-webpart-base" ***!
  \*********************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__642__;

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/compat get default export */
/******/ 	(() => {
/******/ 		// getDefaultExport function for compatibility with non-harmony modules
/******/ 		__webpack_require__.n = (module) => {
/******/ 			var getter = module && module.__esModule ?
/******/ 				() => (module['default']) :
/******/ 				() => (module);
/******/ 			__webpack_require__.d(getter, { a: getter });
/******/ 			return getter;
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/define property getters */
/******/ 	(() => {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = (exports, definition) => {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/global */
/******/ 	(() => {
/******/ 		__webpack_require__.g = (function() {
/******/ 			if (typeof globalThis === 'object') return globalThis;
/******/ 			try {
/******/ 				return this || new Function('return this')();
/******/ 			} catch (e) {
/******/ 				if (typeof window === 'object') return window;
/******/ 			}
/******/ 		})();
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	(() => {
/******/ 		__webpack_require__.o = (obj, prop) => (Object.prototype.hasOwnProperty.call(obj, prop))
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	(() => {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = (exports) => {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	})();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry need to be wrapped in an IIFE because it need to be isolated against other modules in the chunk.
(() => {
/*!****************************************************!*\
  !*** ./lib/webparts/customFaq/CustomFaqWebPart.js ***!
  \****************************************************/
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @microsoft/sp-core-library */ 676);
/* harmony import */ var _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @microsoft/sp-property-pane */ 877);
/* harmony import */ var _microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__);
/* harmony import */ var _microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @microsoft/sp-webpart-base */ 642);
/* harmony import */ var _microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_2___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_2__);
/* harmony import */ var _microsoft_sp_component_base__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @microsoft/sp-component-base */ 962);
/* harmony import */ var _microsoft_sp_component_base__WEBPACK_IMPORTED_MODULE_3___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_component_base__WEBPACK_IMPORTED_MODULE_3__);
/* harmony import */ var _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./CustomFaq.module.scss */ 426);
/* harmony import */ var _services_SharePointService__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ./services/SharePointService */ 666);
var __extends = (undefined && undefined.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();






var CustomFaqWebPart = /** @class */ (function (_super) {
    __extends(CustomFaqWebPart, _super);
    function CustomFaqWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._lists = [];
        _this._columns = [];
        _this._faqItems = [];
        _this._selectedCategory = 'All';
        _this._categories = [];
        _this._searchQuery = '';
        _this._topFaqCategoryKey = '__TOP_FAQ__';
        return _this;
    }
    /**
     * Initialize the web part
     */
    CustomFaqWebPart.prototype.onInit = function () {
        var _this = this;
        return _super.prototype.onInit.call(this).then(function () {
            // Initialize SharePoint service
            _this._spService = new _services_SharePointService__WEBPACK_IMPORTED_MODULE_5__.SharePointService(_this.context);
            // Try to consume the ThemeProvider service for section background support
            try {
                _this._themeProvider = _this.context.serviceScope.consume(_microsoft_sp_component_base__WEBPACK_IMPORTED_MODULE_3__.ThemeProvider.serviceKey);
                if (_this._themeProvider) {
                    _this._themeVariant = _this._themeProvider.tryGetTheme();
                    _this._themeProvider.themeChangedEvent.add(_this, _this._handleThemeChangedEvent);
                }
            }
            catch (error) {
                console.log('ThemeProvider not available:', error);
            }
            if (_this._themeVariant) {
                _this._setCSSVariables(_this._themeVariant);
            }
        }).then(function () {
            return _this._loadLists();
        }).then(function () {
            if (_this.properties.selectedList) {
                return _this._loadColumns().then(function () {
                    return _this._loadFaqItems();
                });
            }
            return Promise.resolve();
        }).catch(function (error) {
            console.error('Error during initialization:', error);
            return Promise.resolve();
        });
    };
    CustomFaqWebPart.prototype._handleThemeChangedEvent = function (args) {
        this._themeVariant = args.theme;
        if (this._themeVariant) {
            this._setCSSVariables(this._themeVariant);
        }
        this.render();
    };
    CustomFaqWebPart.prototype._setCSSVariables = function (theme) {
        if (!this.domElement) {
            return;
        }
        try {
            if (theme.semanticColors) {
                var semanticColors = theme.semanticColors;
                var semanticKeys = Object.keys(semanticColors);
                var i = 0;
                while (i < semanticKeys.length) {
                    var key = semanticKeys[i];
                    var value = semanticColors[key];
                    if (value) {
                        this.domElement.style.setProperty('--' + key, value);
                    }
                    i++;
                }
            }
            if (theme.palette) {
                var palette = theme.palette;
                var paletteKeys = Object.keys(palette);
                var j = 0;
                while (j < paletteKeys.length) {
                    var key = paletteKeys[j];
                    var value = palette[key];
                    if (value) {
                        this.domElement.style.setProperty('--' + key, value);
                    }
                    j++;
                }
            }
        }
        catch (error) {
            console.error('Error setting CSS variables:', error);
        }
    };
    CustomFaqWebPart.prototype._extractCategories = function () {
        var categorySet = {};
        this._categories = [];
        var i = 0;
        while (i < this._faqItems.length) {
            var category = this._faqItems[i].category;
            if (category && !categorySet[category]) {
                categorySet[category] = true;
                this._categories.push(category);
            }
            i++;
        }
        this._categories.sort();
    };
    CustomFaqWebPart.prototype._isTopFaqEnabled = function () {
        return !!(this.properties.topFaqColumn && this.properties.topFaqColumn.trim() !== '');
    };
    CustomFaqWebPart.prototype._hasTopFaqItems = function () {
        var i = 0;
        while (i < this._faqItems.length) {
            if (this._faqItems[i].isTopFaq) {
                return true;
            }
            i++;
        }
        return false;
    };
    CustomFaqWebPart.prototype._ensureValidCategorySelection = function () {
        var validCategories = {
            'All': true
        };
        if (this._isTopFaqEnabled()) {
            validCategories[this._topFaqCategoryKey] = true;
        }
        var i = 0;
        while (i < this._categories.length) {
            validCategories[this._categories[i]] = true;
            i++;
        }
        if (!validCategories[this._selectedCategory]) {
            this._selectedCategory = 'All';
        }
    };
    CustomFaqWebPart.prototype._sanitizeColorInput = function (color) {
        if (!color) {
            return '';
        }
        var trimmed = color.trim();
        if (trimmed === '') {
            return '';
        }
        var hexRegex = /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/;
        var rgbRegex = /^rgba?\(\s*\d{1,3}\s*,\s*\d{1,3}\s*,\s*\d{1,3}(?:\s*,\s*(0|0?\.\d+|1(\.0)?))?\s*\)$/;
        var cssNameRegex = /^[a-zA-Z]+$/;
        if (hexRegex.test(trimmed) || rgbRegex.test(trimmed) || cssNameRegex.test(trimmed)) {
            return trimmed;
        }
        return '';
    };
    CustomFaqWebPart.prototype._getTopFaqColor = function (fallbackColor) {
        var sanitized = this._sanitizeColorInput(this.properties.topFaqColor);
        if (sanitized) {
            return sanitized;
        }
        return fallbackColor;
    };
    CustomFaqWebPart.prototype._getTopFaqBadgePalette = function (color) {
        var badgeBackground = color;
        var badgeBorder = this._hexToRgba(color, 0.35) || 'rgba(0, 0, 0, 0.3)';
        var badgeShadow = this._hexToRgba(color, 0.45) || 'rgba(0, 0, 0, 0.25)';
        var badgeTextColor = this._getContrastingTextColor(color);
        return {
            badgeBackground: badgeBackground,
            badgeTextColor: badgeTextColor,
            badgeBorder: badgeBorder,
            badgeShadow: badgeShadow
        };
    };
    CustomFaqWebPart.prototype._hexToRgba = function (color, alpha) {
        if (!color) {
            return null;
        }
        var trimmed = color.trim();
        if (!trimmed || trimmed.charAt(0) !== '#') {
            return null;
        }
        var hex = trimmed.substring(1);
        if (hex.length === 3) {
            hex = hex.charAt(0) + hex.charAt(0) + hex.charAt(1) + hex.charAt(1) + hex.charAt(2) + hex.charAt(2);
        }
        if (hex.length !== 6 || hex.match(/[^0-9a-fA-F]/)) {
            return null;
        }
        var r = parseInt(hex.substring(0, 2), 16);
        var g = parseInt(hex.substring(2, 4), 16);
        var b = parseInt(hex.substring(4, 6), 16);
        var normalizedAlpha = Math.max(0, Math.min(1, alpha));
        return 'rgba(' + r + ', ' + g + ', ' + b + ', ' + normalizedAlpha + ')';
    };
    CustomFaqWebPart.prototype._parseHexColor = function (color) {
        if (!color || color.charAt(0) !== '#') {
            return null;
        }
        var hex = color.substring(1);
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
    };
    CustomFaqWebPart.prototype._relativeLuminance = function (r, g, b) {
        var srgb = [r, g, b].map(function (value) {
            var normalized = value / 255;
            return normalized <= 0.03928 ? normalized / 12.92 : Math.pow((normalized + 0.055) / 1.055, 2.4);
        });
        return 0.2126 * srgb[0] + 0.7152 * srgb[1] + 0.0722 * srgb[2];
    };
    CustomFaqWebPart.prototype._getContrastingTextColor = function (color) {
        var rgb = this._parseHexColor(color);
        if (!rgb) {
            return '#ffffff';
        }
        var luminance = this._relativeLuminance(rgb.r, rgb.g, rgb.b);
        return luminance > 0.6 ? '#201f1e' : '#ffffff';
    };
    /**
     * Get filtered FAQ items based on selected category and search query
     */
    CustomFaqWebPart.prototype._getFilteredItems = function () {
        var filtered = [];
        // First filter by category
        if (this._selectedCategory === 'All') {
            filtered = this._faqItems.slice();
        }
        else if (this._selectedCategory === this._topFaqCategoryKey) {
            var idx = 0;
            while (idx < this._faqItems.length) {
                if (this._faqItems[idx].isTopFaq) {
                    filtered.push(this._faqItems[idx]);
                }
                idx++;
            }
        }
        else {
            var i = 0;
            while (i < this._faqItems.length) {
                if (this._faqItems[i].category === this._selectedCategory) {
                    filtered.push(this._faqItems[i]);
                }
                i++;
            }
        }
        // Then filter by search query
        if (this._searchQuery && this._searchQuery.trim() !== '') {
            var query = this._searchQuery.toLowerCase().trim();
            var searchResults = [];
            var j = 0;
            while (j < filtered.length) {
                var item = filtered[j];
                var titleMatch = item.title && item.title.toLowerCase().indexOf(query) !== -1;
                var answerMatch = false;
                if (this.properties.searchInAnswers && item.description) {
                    // Strip HTML tags for search using DOM-based sanitization
                    var plainDescription = this._stripHtmlTags(item.description);
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
    };
    /**
     * Get total count of items (before search filter, after category filter)
     */
    CustomFaqWebPart.prototype._getTotalItemsInCategory = function () {
        if (this._selectedCategory === 'All') {
            return this._faqItems.length;
        }
        if (this._selectedCategory === this._topFaqCategoryKey) {
            var topCount = 0;
            var idx = 0;
            while (idx < this._faqItems.length) {
                if (this._faqItems[idx].isTopFaq) {
                    topCount++;
                }
                idx++;
            }
            return topCount;
        }
        var count = 0;
        var i = 0;
        while (i < this._faqItems.length) {
            if (this._faqItems[i].category === this._selectedCategory) {
                count++;
            }
            i++;
        }
        return count;
    };
    /**
     * Highlight search matches in text
     */
    CustomFaqWebPart.prototype._highlightText = function (text, isHtml) {
        if (!this._searchQuery || this._searchQuery.trim() === '' || !this.properties.highlightSearchMatches) {
            return text;
        }
        var query = this._searchQuery.trim();
        if (isHtml) {
            // For HTML content, use DOM-based approach for safe highlighting
            return this._highlightInHtml(text, query);
        }
        else {
            // For plain text, simple replacement is safe
            var regex = new RegExp('(' + this._escapeRegExp(query) + ')', 'gi');
            return text.replace(regex, '<mark class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].searchHighlight + '">$1</mark>');
        }
    };
    /**
     * Safely highlight text within HTML content using DOM parsing
     */
    CustomFaqWebPart.prototype._highlightInHtml = function (html, query) {
        try {
            var parser = new DOMParser();
            var doc = parser.parseFromString('<div>' + html + '</div>', 'text/html');
            var container = doc.body.firstChild;
            if (!container) {
                return html;
            }
            // Walk through text nodes only
            var walker = doc.createTreeWalker(container, NodeFilter.SHOW_TEXT, null);
            var textNodes = [];
            var node = void 0;
            while ((node = walker.nextNode())) {
                textNodes.push(node);
            }
            var lowerQuery = query.toLowerCase();
            var i = 0;
            while (i < textNodes.length) {
                var textNode = textNodes[i];
                var text = textNode.textContent || '';
                var lowerText = text.toLowerCase();
                var index = lowerText.indexOf(lowerQuery);
                if (index !== -1 && textNode.parentNode) {
                    var before = text.substring(0, index);
                    var match = text.substring(index, index + query.length);
                    var after = text.substring(index + query.length);
                    var fragment = doc.createDocumentFragment();
                    if (before) {
                        fragment.appendChild(doc.createTextNode(before));
                    }
                    var mark = doc.createElement('mark');
                    mark.className = _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].searchHighlight;
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
        }
        catch (error) {
            // Fallback: return original HTML if parsing fails
            console.warn('HTML highlighting failed:', error);
            return html;
        }
    };
    /**
     * Securely strip HTML tags using DOM-based parsing
     * Addresses CodeQL js/incomplete-multi-character-sanitization
     * and js/redos (ReDoS vulnerability)
     */
    CustomFaqWebPart.prototype._stripHtmlTags = function (input) {
        if (!input) {
            return '';
        }
        try {
            // Use DOMParser for safe HTML stripping - no regex vulnerability
            var parser = new DOMParser();
            var doc = parser.parseFromString('<div>' + input + '</div>', 'text/html');
            var container = doc.body.firstChild;
            if (!container) {
                return this._fallbackStripTags(input);
            }
            // Get text content only (strips all HTML)
            var result = container.textContent || '';
            // Decode common HTML entities
            return this._decodeHtmlEntities(result);
        }
        catch (error) {
            // Fallback if DOMParser fails
            console.warn('DOMParser failed, using fallback:', error);
            return this._fallbackStripTags(input);
        }
    };
    /**
     * Fallback tag stripper using character-by-character parsing
     * Safe from ReDoS as it doesn't use regex with backtracking
     */
    CustomFaqWebPart.prototype._fallbackStripTags = function (input) {
        var result = '';
        var inTag = false;
        var chars = input.split('');
        var idx = 0;
        var len = chars.length;
        while (idx < len) {
            var char = chars[idx];
            if (char === '<') {
                inTag = true;
            }
            else if (char === '>') {
                inTag = false;
            }
            else if (!inTag) {
                result += char;
            }
            idx++;
        }
        return this._decodeHtmlEntities(result);
    };
    /**
     * Decode common HTML entities using split/join (ReDoS safe)
     */
    CustomFaqWebPart.prototype._decodeHtmlEntities = function (text) {
        // Use split/join instead of regex to avoid ReDoS
        // Process each entity directly to avoid for loop lint warning
        var result = text;
        result = result.split('&nbsp;').join(' ');
        result = result.split('&amp;').join('&');
        result = result.split('&lt;').join('<');
        result = result.split('&gt;').join('>');
        result = result.split('&quot;').join('"');
        result = result.split('&#39;').join("'");
        result = result.split('&#x27;').join("'");
        result = result.split('&apos;').join("'");
        return result;
    };
    /**
     * Escape special regex characters
     */
    CustomFaqWebPart.prototype._escapeRegExp = function (text) {
        return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    };
    CustomFaqWebPart.prototype._getFontSize = function (size, defaultSize) {
        return (size || defaultSize) + 'px';
    };
    /**
     * Render the web part
     */
    CustomFaqWebPart.prototype.render = function () {
        try {
            var semanticColors = this._themeVariant ? this._themeVariant.semanticColors : undefined;
            var palette = this._themeVariant ? this._themeVariant.palette : undefined;
            var headerBgColor = (palette && palette.themePrimary) ? palette.themePrimary : '#0078d4';
            var headerBgColorDark = (palette && palette.themeDark) ? palette.themeDark : '#005a9e';
            var bodyBackground = (semanticColors && semanticColors.bodyBackground) ? semanticColors.bodyBackground : '#ffffff';
            var bodyText = (semanticColors && semanticColors.bodyText) ? semanticColors.bodyText : '#323130';
            var bodySubtext = (semanticColors && semanticColors.bodySubtext) ? semanticColors.bodySubtext : '#605e5c';
            var linkColor = (semanticColors && semanticColors.link) ? semanticColors.link : ((palette && palette.themePrimary) ? palette.themePrimary : '#0078d4');
            var bodyDivider = (semanticColors && semanticColors.bodyDivider) ? semanticColors.bodyDivider : '#edebe9';
            var listItemBackgroundHovered = (semanticColors && semanticColors.listItemBackgroundHovered) ? semanticColors.listItemBackgroundHovered : '#f3f2f1';
            var cardBackground = (semanticColors && semanticColors.cardStandoutBackground) ? semanticColors.cardStandoutBackground : bodyBackground;
            var neutralLighter = (palette && palette.neutralLighter) ? palette.neutralLighter : '#f4f4f4';
            var neutralTertiary = (palette && palette.neutralTertiary) ? palette.neutralTertiary : '#a6a6a6';
            var neutralLight = (palette && palette.neutralLight) ? palette.neutralLight : '#eaeaea';
            var inputBackground = (semanticColors && semanticColors.inputBackground) ? semanticColors.inputBackground : '#ffffff';
            var inputBorder = (semanticColors && semanticColors.inputBorder) ? semanticColors.inputBorder : '#8a8886';
            var webpartTitleFontSize = this._getFontSize(this.properties.webpartTitleFontSize, '24');
            var webpartDescriptionFontSize = this._getFontSize(this.properties.webpartDescriptionFontSize, '14');
            var faqTitleFontSize = this._getFontSize(this.properties.faqTitleFontSize, '16');
            var faqDescriptionFontSize = this._getFontSize(this.properties.faqDescriptionFontSize, '14');
            var searchPlaceholder = this.properties.searchPlaceholder || 'Search FAQs...';
            var faqItemsHtml = '';
            var categoryTabsHtml = '';
            var searchHtml = '';
            var searchInHeaderHtml = '';
            var resultsInfoHtml = '';
            // Build search input HTML
            var searchInputHtml = '<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].searchContainer + '">' +
                '<svg class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].searchIcon + '" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
                '<circle cx="11" cy="11" r="8"/>' +
                '<path d="M21 21l-4.35-4.35"/>' +
                '</svg>' +
                '<input type="text" class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].searchInput + '" ' +
                'placeholder="' + this._escapeHtml(searchPlaceholder) + '" ' +
                'value="' + this._escapeHtml(this._searchQuery) + '" ' +
                'style="background-color: ' + inputBackground + '; border-color: ' + inputBorder + '; color: ' + bodyText + ';">' +
                (this._searchQuery ? '<button class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].clearButton + '" title="Clear search" style="color: ' + bodySubtext + ';">' +
                    '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
                    '<line x1="18" y1="6" x2="6" y2="18"/>' +
                    '<line x1="6" y1="6" x2="18" y2="18"/>' +
                    '</svg>' +
                    '</button>' : '') +
                '</div>';
            if (!this.properties.selectedList) {
                faqItemsHtml = '<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].emptyState + '" style="color: ' + bodySubtext + ';">' +
                    '<svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="' + neutralTertiary + '" stroke-width="1.5">' +
                    '<circle cx="12" cy="12" r="10"/>' +
                    '<path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"/>' +
                    '<line x1="12" y1="17" x2="12.01" y2="17"/>' +
                    '</svg>' +
                    '<p>Please configure the web part by selecting a list from the property pane.</p>' +
                    '</div>';
            }
            else {
                // Build search section (below header)
                if (this.properties.enableSearch && !this.properties.searchInHeader) {
                    searchHtml = '<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].searchSection + '" style="background-color: ' + neutralLighter + '; border-bottom-color: ' + bodyDivider + ';">' +
                        searchInputHtml +
                        '</div>';
                }
                // Build search in header
                if (this.properties.enableSearch && this.properties.searchInHeader) {
                    searchInHeaderHtml = '<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].searchInHeader + '">' + searchInputHtml + '</div>';
                }
                // Build category tabs (All + optional Top FAQ + categories)
                var hasCategoryTabs = !!(this.properties.categoryColumn && this._categories.length > 0);
                var showTopFaqTab = this._isTopFaqEnabled();
                if (hasCategoryTabs || showTopFaqTab) {
                    var tabsArray = [];
                    var allActiveClass = this._selectedCategory === 'All' ? ' ' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].activeTab : '';
                    tabsArray.push('<button class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].categoryTab + allActiveClass + '" ' +
                        'data-category="All" ' +
                        'style="color: ' + (this._selectedCategory === 'All' ? headerBgColor : bodyText) + '; ' +
                        'border-bottom-color: ' + (this._selectedCategory === 'All' ? headerBgColor : 'transparent') + ';">' +
                        'All' +
                        '</button>');
                    if (showTopFaqTab) {
                        var isTopActive = this._selectedCategory === this._topFaqCategoryKey;
                        var topActiveClass = isTopActive ? ' ' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].activeTab : '';
                        var topFaqColor = this._getTopFaqColor(headerBgColor);
                        var topBadgePalette = this._getTopFaqBadgePalette(topFaqColor);
                        var badgeStyle = ' style="--topFaqBadgeBackground: ' + this._escapeHtml(topBadgePalette.badgeBackground) +
                            '; --topFaqBadgeColor: ' + this._escapeHtml(topBadgePalette.badgeTextColor) +
                            '; --topFaqBadgeBorder: ' + this._escapeHtml(topBadgePalette.badgeBorder) +
                            '; --topFaqBadgeShadow: ' + this._escapeHtml(topBadgePalette.badgeShadow) + ';"';
                        tabsArray.push('<button class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].categoryTab + topActiveClass + '" ' +
                            'data-category="' + this._topFaqCategoryKey + '" ' +
                            'style="color: ' + (isTopActive ? headerBgColor : bodyText) + '; ' +
                            'border-bottom-color: ' + (isTopActive ? headerBgColor : 'transparent') + ';">' +
                            '<span class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].topFaqTab + '">' +
                            '<span class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].topFaqBadge + '"' + badgeStyle + '>⭐</span>' +
                            '<span class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].topFaqBadgeLabel + '">Top FAQ</span>' +
                            '</span>' +
                            '</button>');
                    }
                    if (hasCategoryTabs) {
                        var catIdx = 0;
                        while (catIdx < this._categories.length) {
                            var category = this._categories[catIdx];
                            var isActive = this._selectedCategory === category;
                            var activeClass = isActive ? ' ' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].activeTab : '';
                            tabsArray.push('<button class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].categoryTab + activeClass + '" ' +
                                'data-category="' + this._escapeHtml(category) + '" ' +
                                'style="color: ' + (isActive ? headerBgColor : bodyText) + '; ' +
                                'border-bottom-color: ' + (isActive ? headerBgColor : 'transparent') + ';">' +
                                this._escapeHtml(category) +
                                '</button>');
                            catIdx++;
                        }
                    }
                    categoryTabsHtml = '<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].categoryTabs + '" style="border-bottom-color: ' + neutralLight + ';">' +
                        tabsArray.join('') +
                        '</div>';
                }
                // Get filtered items
                var filteredItems = this._getFilteredItems();
                var totalItems = this._getTotalItemsInCategory();
                // Build results info
                if (this.properties.enableSearch && this.properties.showResultsCount && this._searchQuery && this._searchQuery.trim() !== '') {
                    resultsInfoHtml = '<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].resultsInfo + '" style="color: ' + bodySubtext + ';">' +
                        'Showing <strong style="color: ' + bodyText + ';">' + filteredItems.length + '</strong> of ' +
                        '<strong style="color: ' + bodyText + ';">' + totalItems + '</strong> FAQs matching ' +
                        '"<strong style="color: ' + bodyText + ';">' + this._escapeHtml(this._searchQuery) + '</strong>"' +
                        '</div>';
                }
                if (filteredItems.length === 0) {
                    var noResultsMessage = this._searchQuery && this._searchQuery.trim() !== ''
                        ? 'No FAQs found matching "' + this._escapeHtml(this._searchQuery) + '". Try different keywords or clear the search.'
                        : 'No FAQ items found' + (this._selectedCategory !== 'All' ? ' in this category' : '') + '.';
                    faqItemsHtml = '<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].emptyState + '" style="color: ' + bodySubtext + ';">' +
                        '<svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="' + neutralTertiary + '" stroke-width="1.5">' +
                        (this._searchQuery ?
                            '<circle cx="11" cy="11" r="8"/><path d="M21 21l-4.35-4.35"/><line x1="8" y1="11" x2="14" y2="11"/>' :
                            '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/>') +
                        '</svg>' +
                        '<p>' + noResultsMessage + '</p>' +
                        '</div>';
                }
                else {
                    var itemsHtmlArray = [];
                    var index = 0;
                    while (index < filteredItems.length) {
                        var item = filteredItems[index];
                        var attachmentsHtml = '';
                        if (item.attachments && item.attachments.length > 0) {
                            var attachmentLinksArray = [];
                            var attIdx = 0;
                            while (attIdx < item.attachments.length) {
                                var att = item.attachments[attIdx];
                                attachmentLinksArray.push('<a href="' + att.url + '" target="_blank" rel="noopener noreferrer" class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].attachmentLink + '" style="color: ' + linkColor + ';">' +
                                    '<span class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].attachmentIcon + '">' +
                                    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
                                    '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>' +
                                    '<polyline points="14 2 14 8 20 8"/>' +
                                    '</svg>' +
                                    '</span>' +
                                    this._escapeHtml(att.fileName) +
                                    '</a>');
                                attIdx++;
                            }
                            attachmentsHtml = '<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].attachments + '" style="background-color: ' + neutralLighter + ';">' +
                                '<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].attachmentsLabel + '" style="color: ' + neutralTertiary + ';">' +
                                '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
                                '<path d="M21.44 11.05l-9.19 9.19a6 6 0 0 1-8.49-8.49l9.19-9.19a4 4 0 0 1 5.66 5.66l-9.2 9.19a2 2 0 0 1-2.83-2.83l8.49-8.48"/>' +
                                '</svg>' +
                                ' Attachments' +
                                '</div>' +
                                attachmentLinksArray.join('') +
                                '</div>';
                        }
                        // Apply highlighting to title
                        var highlightedTitle = this._highlightText(this._escapeHtml(item.title), false);
                        // Apply highlighting to description
                        var formattedDescription = this._formatDescription(item.description);
                        var highlightedDescription = this.properties.searchInAnswers
                            ? this._highlightText(formattedDescription, formattedDescription.indexOf('<') !== -1)
                            : formattedDescription;
                        itemsHtmlArray.push('<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqItem + '" data-index="' + index + '">' +
                            '<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqQuestion + '" style="border-bottom-color: ' + bodyDivider + ';" data-hover-bg="' + listItemBackgroundHovered + '">' +
                            '<span class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqQuestionText + '" style="color: ' + bodyText + '; font-size: ' + faqTitleFontSize + ';">' +
                            highlightedTitle +
                            '</span>' +
                            '<span class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqChevron + '" style="color: ' + bodySubtext + ';">' +
                            '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
                            '<path d="M6 9l6 6 6-6"/>' +
                            '</svg>' +
                            '</span>' +
                            '</div>' +
                            '<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqAnswer + '">' +
                            '<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqAnswerContent + '" style="color: ' + bodySubtext + '; font-size: ' + faqDescriptionFontSize + ';">' +
                            highlightedDescription +
                            '</div>' +
                            attachmentsHtml +
                            '</div>' +
                            '</div>');
                        index++;
                    }
                    faqItemsHtml = itemsHtmlArray.join('');
                }
            }
            // Build header section
            var headerHtml = '';
            if (this.properties.webpartTitle || this.properties.webpartDescription || searchInHeaderHtml) {
                headerHtml = '<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqHeader + '" style="background: linear-gradient(135deg, ' + headerBgColor + ' 0%, ' + headerBgColorDark + ' 100%);">';
                if (this.properties.webpartTitle) {
                    headerHtml += '<h2 class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqTitle + '" style="font-size: ' + webpartTitleFontSize + ';">' + this._escapeHtml(this.properties.webpartTitle) + '</h2>';
                }
                if (this.properties.webpartDescription) {
                    headerHtml += '<p class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqDescription + '" style="font-size: ' + webpartDescriptionFontSize + ';">' + this._escapeHtml(this.properties.webpartDescription) + '</p>';
                }
                headerHtml += searchInHeaderHtml;
                headerHtml += '</div>';
            }
            this.domElement.innerHTML = '<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].customFaq + '" style="background-color: ' + cardBackground + ';">' +
                headerHtml +
                searchHtml +
                resultsInfoHtml +
                categoryTabsHtml +
                '<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqList + '">' +
                faqItemsHtml +
                '</div>' +
                '</div>';
            this._attachEventListeners();
            this._attachTabEventListeners();
            this._attachSearchEventListeners();
        }
        catch (error) {
            console.error('Error during render:', error);
            this.domElement.innerHTML = '<div style="padding: 20px; color: red;">Error rendering web part. Please check console for details.</div>';
        }
    };
    /**
     * Attach search event listeners
     */
    CustomFaqWebPart.prototype._attachSearchEventListeners = function () {
        var searchInput = this.domElement.querySelector('.' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].searchInput);
        var clearButton = this.domElement.querySelector('.' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].clearButton);
        var self = this;
        if (searchInput) {
            // Debounce search input
            var debounceTimer_1;
            searchInput.addEventListener('input', function () {
                if (debounceTimer_1) {
                    clearTimeout(debounceTimer_1);
                }
                debounceTimer_1 = setTimeout(function () {
                    self._searchQuery = searchInput.value;
                    self.render();
                }, 300);
            });
            // Handle Escape key
            searchInput.addEventListener('keydown', function (e) {
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
            clearButton.addEventListener('click', function () {
                self._searchQuery = '';
                self.render();
            });
        }
    };
    CustomFaqWebPart.prototype._attachTabEventListeners = function () {
        var tabs = this.domElement.querySelectorAll('.' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].categoryTab);
        var self = this;
        var tabIdx = 0;
        while (tabIdx < tabs.length) {
            var tab = tabs[tabIdx];
            (function (t) {
                t.addEventListener('click', function () {
                    var category = t.dataset.category;
                    if (category) {
                        self._selectedCategory = category;
                        self.render();
                    }
                });
            })(tab);
            tabIdx++;
        }
    };
    CustomFaqWebPart.prototype._attachEventListeners = function () {
        var faqItems = this.domElement.querySelectorAll('.' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqItem);
        var questions = this.domElement.querySelectorAll('.' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqQuestion);
        var self = this;
        var index = 0;
        while (index < questions.length) {
            var questionElement = questions[index];
            (function (qEl, idx) {
                qEl.addEventListener('mouseenter', function () {
                    var hoverBg = qEl.dataset.hoverBg;
                    if (hoverBg) {
                        qEl.style.backgroundColor = hoverBg;
                    }
                });
                qEl.addEventListener('mouseleave', function () {
                    qEl.style.backgroundColor = '';
                });
                qEl.addEventListener('click', function () {
                    var faqItem = faqItems[idx];
                    if (!self.properties.allowMultipleExpanded) {
                        var i = 0;
                        while (i < faqItems.length) {
                            if (i !== idx && faqItems[i].classList.contains(_CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].expanded)) {
                                faqItems[i].classList.remove(_CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].expanded);
                            }
                            i++;
                        }
                    }
                    if (faqItem.classList.contains(_CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].expanded)) {
                        faqItem.classList.remove(_CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].expanded);
                    }
                    else {
                        faqItem.classList.add(_CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].expanded);
                    }
                });
            })(questionElement, index);
            index++;
        }
    };
    CustomFaqWebPart.prototype._escapeHtml = function (text) {
        if (!text) {
            return '';
        }
        var div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    };
    CustomFaqWebPart.prototype._formatDescription = function (description) {
        if (!description) {
            return '';
        }
        // Check if content appears to contain HTML
        var hasOpenTag = description.indexOf('<') !== -1;
        var hasCloseTag = description.indexOf('>') !== -1;
        if (hasOpenTag && hasCloseTag) {
            // Content has HTML, return as-is (will be rendered as HTML)
            return description;
        }
        // Plain text: escape HTML entities and convert newlines
        return this._escapeHtml(description).split('\n').join('<br>');
    };
    CustomFaqWebPart.prototype._loadLists = function () {
        var _this = this;
        return this._spService.getLists()
            .then(function (lists) {
            _this._lists = lists;
        })
            .catch(function (error) {
            console.error('Error loading lists:', error);
            _this._lists = [];
        });
    };
    CustomFaqWebPart.prototype._loadColumns = function () {
        var _this = this;
        if (!this.properties.selectedList) {
            this._columns = [];
            return Promise.resolve();
        }
        return this._spService.getListColumns(this.properties.selectedList)
            .then(function (columns) {
            _this._columns = columns;
        })
            .catch(function (error) {
            console.error('Error loading columns:', error);
            _this._columns = [];
        });
    };
    CustomFaqWebPart.prototype._loadFaqItems = function () {
        var _this = this;
        if (!this.properties.selectedList || !this.properties.titleColumn || !this.properties.descriptionColumn) {
            this._faqItems = [];
            return Promise.resolve();
        }
        return this._spService.getListItems(this.properties.selectedList, this.properties.titleColumn, this.properties.descriptionColumn, this.properties.categoryColumn || undefined, this.properties.topFaqColumn || undefined)
            .then(function (items) {
            _this._faqItems = items;
            _this._extractCategories();
            var hasTopFaqItems = _this._isTopFaqEnabled() && _this._hasTopFaqItems();
            if (hasTopFaqItems && _this._selectedCategory === 'All') {
                _this._selectedCategory = _this._topFaqCategoryKey;
            }
            _this._ensureValidCategorySelection();
        })
            .catch(function (error) {
            console.error('Error loading FAQ items:', error);
            _this._faqItems = [];
            _this._categories = [];
        });
    };
    Object.defineProperty(CustomFaqWebPart.prototype, "dataVersion", {
        get: function () {
            return _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_0__.Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    CustomFaqWebPart.prototype.onPropertyPaneFieldChanged = function (propertyPath, oldValue, newValue) {
        var self = this;
        if (propertyPath === 'selectedList' && newValue !== oldValue) {
            this.properties.titleColumn = '';
            this.properties.descriptionColumn = '';
            this.properties.categoryColumn = '';
            this.properties.topFaqColumn = '';
            this._faqItems = [];
            this._categories = [];
            this._selectedCategory = 'All';
            this._searchQuery = '';
            this._loadColumns().then(function () {
                self.context.propertyPane.refresh();
                self.render();
            });
        }
        else if ((propertyPath === 'titleColumn' || propertyPath === 'descriptionColumn' || propertyPath === 'categoryColumn' || propertyPath === 'topFaqColumn') && newValue !== oldValue) {
            this._loadFaqItems().then(function () {
                self.render();
            });
        }
        else {
            this.render();
        }
    };
    CustomFaqWebPart.prototype._getListOptions = function () {
        var options = [
            { key: '', text: '-- Select a list --' }
        ];
        var i = 0;
        while (i < this._lists.length) {
            options.push({ key: this._lists[i].id, text: this._lists[i].title });
            i++;
        }
        return options;
    };
    CustomFaqWebPart.prototype._getTitleColumnOptions = function () {
        var options = [
            { key: '', text: '-- Select a column --' }
        ];
        var i = 0;
        while (i < this._columns.length) {
            var col = this._columns[i];
            if (col.type === 'Text' || col.type === 'Note') {
                options.push({ key: col.internalName, text: col.title });
            }
            i++;
        }
        return options;
    };
    CustomFaqWebPart.prototype._getDescriptionColumnOptions = function () {
        var options = [
            { key: '', text: '-- Select a column --' }
        ];
        var i = 0;
        while (i < this._columns.length) {
            var col = this._columns[i];
            if (col.type === 'Text' || col.type === 'Note') {
                var typeLabel = col.type === 'Note' ? 'Multi-line' : 'Single-line';
                options.push({ key: col.internalName, text: col.title + ' (' + typeLabel + ')' });
            }
            i++;
        }
        return options;
    };
    CustomFaqWebPart.prototype._getCategoryColumnOptions = function () {
        var options = [
            { key: '', text: '-- No category (disable tabs) --' }
        ];
        var i = 0;
        while (i < this._columns.length) {
            var col = this._columns[i];
            if (col.type === 'Text' || col.type === 'Choice') {
                options.push({ key: col.internalName, text: col.title });
            }
            i++;
        }
        return options;
    };
    CustomFaqWebPart.prototype._getTopFaqColumnOptions = function () {
        var options = [
            { key: '', text: '-- No Top FAQ column --' }
        ];
        var i = 0;
        while (i < this._columns.length) {
            var col = this._columns[i];
            if (col.type === 'Boolean' || col.type === 'Choice' || col.type === 'Text') {
                options.push({ key: col.internalName, text: col.title });
            }
            i++;
        }
        return options;
    };
    CustomFaqWebPart.prototype._getFontSizeOptions = function () {
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
    };
    CustomFaqWebPart.prototype.getPropertyPaneConfiguration = function () {
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
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneTextField)('webpartTitle', {
                                    label: 'Webpart Title',
                                    placeholder: 'Enter a title for the FAQ section'
                                }),
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneDropdown)('webpartTitleFontSize', {
                                    label: 'Webpart Title Font Size',
                                    options: this._getFontSizeOptions(),
                                    selectedKey: this.properties.webpartTitleFontSize || '24'
                                }),
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneTextField)('webpartDescription', {
                                    label: 'Webpart Description',
                                    placeholder: 'Enter a description',
                                    multiline: true,
                                    rows: 3
                                }),
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneDropdown)('webpartDescriptionFontSize', {
                                    label: 'Webpart Description Font Size',
                                    options: this._getFontSizeOptions(),
                                    selectedKey: this.properties.webpartDescriptionFontSize || '14'
                                }),
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneTextField)('topFaqColor', {
                                    label: 'Top FAQ Accent Color',
                                    placeholder: '#0078d4 or red',
                                    description: 'Leave blank to inherit the theme color from the page.'
                                })
                            ]
                        },
                        {
                            groupName: 'FAQ Item Settings',
                            groupFields: [
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneDropdown)('faqTitleFontSize', {
                                    label: 'FAQ Question Font Size',
                                    options: this._getFontSizeOptions(),
                                    selectedKey: this.properties.faqTitleFontSize || '16'
                                }),
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneDropdown)('faqDescriptionFontSize', {
                                    label: 'FAQ Answer Font Size',
                                    options: this._getFontSizeOptions(),
                                    selectedKey: this.properties.faqDescriptionFontSize || '14'
                                })
                            ]
                        },
                        {
                            groupName: 'Search Settings',
                            groupFields: [
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneToggle)('enableSearch', {
                                    label: 'Enable Search',
                                    onText: 'Yes',
                                    offText: 'No'
                                }),
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneToggle)('searchInHeader', {
                                    label: 'Search in Header',
                                    onText: 'Yes (in header)',
                                    offText: 'No (below header)',
                                    disabled: !this.properties.enableSearch
                                }),
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneToggle)('highlightSearchMatches', {
                                    label: 'Highlight Search Matches',
                                    onText: 'Yes',
                                    offText: 'No',
                                    disabled: !this.properties.enableSearch
                                }),
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneToggle)('searchInAnswers', {
                                    label: 'Search in Answers',
                                    onText: 'Yes (titles & answers)',
                                    offText: 'No (titles only)',
                                    disabled: !this.properties.enableSearch
                                }),
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneToggle)('showResultsCount', {
                                    label: 'Show Results Count',
                                    onText: 'Yes',
                                    offText: 'No',
                                    disabled: !this.properties.enableSearch
                                }),
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneTextField)('searchPlaceholder', {
                                    label: 'Search Placeholder Text',
                                    placeholder: 'Search FAQs...',
                                    disabled: !this.properties.enableSearch
                                })
                            ]
                        },
                        {
                            groupName: 'Data Source',
                            groupFields: [
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneDropdown)('selectedList', {
                                    label: 'Select List',
                                    options: this._getListOptions()
                                }),
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneDropdown)('titleColumn', {
                                    label: 'Title Column',
                                    options: this._getTitleColumnOptions(),
                                    disabled: !this.properties.selectedList
                                }),
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneDropdown)('descriptionColumn', {
                                    label: 'Description Column',
                                    options: this._getDescriptionColumnOptions(),
                                    disabled: !this.properties.selectedList
                                }),
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneDropdown)('categoryColumn', {
                                    label: 'Category Column (for tabs)',
                                    options: this._getCategoryColumnOptions(),
                                    disabled: !this.properties.selectedList
                                }),
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneDropdown)('topFaqColumn', {
                                    label: 'Top FAQ Column (Yes/No)',
                                    options: this._getTopFaqColumnOptions(),
                                    disabled: !this.properties.selectedList
                                })
                            ]
                        },
                        {
                            groupName: 'Accordion Behavior',
                            groupFields: [
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneToggle)('allowMultipleExpanded', {
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
    };
    return CustomFaqWebPart;
}(_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_2__.BaseClientSideWebPart));
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (CustomFaqWebPart);

})();

/******/ 	return __webpack_exports__;
/******/ })()
;
});;
//# sourceMappingURL=custom-faq-web-part.js.map