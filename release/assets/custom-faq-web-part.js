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


_node_modules_microsoft_sp_css_loader_node_modules_microsoft_load_themed_styles_lib_es6_index_js__WEBPACK_IMPORTED_MODULE_0__.loadStyles(".customFaq_f7a6a6e3{background-color:\"[theme:bodyBackground, default: #ffffff]\";background-color:var(--bodyBackground,#fff);border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,.08);font-family:Segoe UI,-apple-system,BlinkMacSystemFont,Roboto,Helvetica Neue,sans-serif;overflow:hidden}.faqHeader_f7a6a6e3{background:\"[theme:themePrimary, default: #0078d4]\";background:var(--themePrimary,#0078d4);color:\"[theme:white, default: #ffffff]\";color:var(--white,#fff);padding:24px 28px}.faqTitle_f7a6a6e3{font-weight:600;line-height:1.3;margin:0 0 4px}.faqDescription_f7a6a6e3{line-height:1.4;margin:0;opacity:.9}.searchSection_f7a6a6e3{background-color:\"[theme:neutralLighter, default: #f4f4f4]\";background-color:var(--neutralLighter,#f4f4f4);border-bottom:1px solid;border-bottom-color:var(--bodyDivider,#edebe9);padding:16px 20px}.searchContainer_f7a6a6e3{max-width:100%;position:relative}.searchIcon_f7a6a6e3{color:\"[theme:bodySubtext, default: #605e5c]\";color:var(--bodySubtext,#605e5c);height:16px;left:12px;pointer-events:none;position:absolute;top:50%;transform:translateY(-50%);width:16px}.searchInput_f7a6a6e3{background-color:\"[theme:inputBackground, default: #ffffff]\";background-color:var(--inputBackground,#fff);border:1px solid;border-color:var(--inputBorder,#8a8886);border-radius:4px;color:\"[theme:bodyText, default: #323130]\";color:var(--bodyText,#323130);font-family:inherit;font-size:14px;padding:10px 40px;transition:border-color .2s,box-shadow .2s;width:100%}.searchInput:-ms-input-placeholder{color:\"[theme:inputPlaceholderText, default: #a19f9d]\";color:var(--inputPlaceholderText,#a19f9d)}.searchInput_f7a6a6e3:-ms-input-placeholder{color:\"[theme:inputPlaceholderText, default: #a19f9d]\";color:var(--inputPlaceholderText,#a19f9d)}.searchInput_f7a6a6e3::placeholder{color:\"[theme:inputPlaceholderText, default: #a19f9d]\";color:var(--inputPlaceholderText,#a19f9d)}.searchInput_f7a6a6e3:focus{border-color:\"[theme:themePrimary, default: #0078d4]\";border-color:var(--themePrimary,#0078d4);box-shadow:0 0 0 1px var(--themePrimary,#0078d4);outline:0}.clearButton_f7a6a6e3{align-items:center;background:0 0;border:none;border-radius:2px;color:\"[theme:bodySubtext, default: #605e5c]\";color:var(--bodySubtext,#605e5c);cursor:pointer;display:flex;justify-content:center;padding:4px;position:absolute;right:8px;top:50%;transform:translateY(-50%)}.clearButton_f7a6a6e3:hover{background-color:\"[theme:neutralLight, default: #edebe9]\";background-color:var(--neutralLight,#edebe9)}.clearButton_f7a6a6e3:focus{outline:2px solid;outline-color:var(--focusBorder,#0078d4);outline-offset:1px}.searchInHeader_f7a6a6e3{margin-top:16px}.searchInHeader_f7a6a6e3 .searchContainer_f7a6a6e3{max-width:100%}.searchInHeader_f7a6a6e3 .searchInput_f7a6a6e3{background-color:hsla(0,0%,100%,.15);border-color:hsla(0,0%,100%,.3);color:#fff}.searchInHeader .searchInput:-ms-input-placeholder{color:hsla(0,0%,100%,.7)}.searchInHeader_f7a6a6e3 .searchInput_f7a6a6e3:-ms-input-placeholder{color:hsla(0,0%,100%,.7)}.searchInHeader_f7a6a6e3 .searchInput_f7a6a6e3::placeholder{color:hsla(0,0%,100%,.7)}.searchInHeader_f7a6a6e3 .searchInput_f7a6a6e3:focus{background-color:#fff;border-color:#fff;box-shadow:none;color:\"[theme:bodyText, default: #323130]\";color:var(--bodyText,#323130)}.searchInHeader .searchInput:focus:-ms-input-placeholder{color:\"[theme:inputPlaceholderText, default: #a19f9d]\";color:var(--inputPlaceholderText,#a19f9d)}.searchInHeader_f7a6a6e3 .searchInput_f7a6a6e3:focus:-ms-input-placeholder{color:\"[theme:inputPlaceholderText, default: #a19f9d]\";color:var(--inputPlaceholderText,#a19f9d)}.searchInHeader_f7a6a6e3 .searchInput_f7a6a6e3:focus::placeholder{color:\"[theme:inputPlaceholderText, default: #a19f9d]\";color:var(--inputPlaceholderText,#a19f9d)}.searchInHeader_f7a6a6e3 .searchIcon_f7a6a6e3{color:hsla(0,0%,100%,.7)}.searchInHeader_f7a6a6e3 .searchContainer_f7a6a6e3:focus-within .searchIcon_f7a6a6e3,.searchInHeader_f7a6a6e3 .searchInput_f7a6a6e3:focus~.searchIcon_f7a6a6e3{color:\"[theme:bodySubtext, default: #605e5c]\";color:var(--bodySubtext,#605e5c)}.searchInHeader_f7a6a6e3 .clearButton_f7a6a6e3{color:hsla(0,0%,100%,.7)}.searchInHeader_f7a6a6e3 .clearButton_f7a6a6e3:hover{background-color:hsla(0,0%,100%,.2)}.searchInHeader_f7a6a6e3 .searchContainer_f7a6a6e3:focus-within .clearButton_f7a6a6e3,.searchInHeader_f7a6a6e3 .searchInput_f7a6a6e3:focus~.clearButton_f7a6a6e3{color:\"[theme:bodySubtext, default: #605e5c]\";color:var(--bodySubtext,#605e5c)}.searchInHeader_f7a6a6e3 .searchContainer_f7a6a6e3:focus-within .clearButton_f7a6a6e3:hover,.searchInHeader_f7a6a6e3 .searchInput_f7a6a6e3:focus~.clearButton_f7a6a6e3:hover{background-color:\"[theme:neutralLight, default: #edebe9]\";background-color:var(--neutralLight,#edebe9)}.resultsInfo_f7a6a6e3{background-color:\"[theme:neutralLighter, default: #f4f4f4]\";background-color:var(--neutralLighter,#f4f4f4);border-bottom:1px solid;border-bottom-color:var(--bodyDivider,#edebe9);color:\"[theme:bodySubtext, default: #605e5c]\";color:var(--bodySubtext,#605e5c);font-size:12px;padding:8px 20px}.resultsInfo_f7a6a6e3 strong{color:\"[theme:bodyText, default: #323130]\";color:var(--bodyText,#323130)}.searchHighlight_f7a6a6e3{background-color:#fff3cd;border-radius:2px;padding:1px 2px}@media (forced-colors:active){.searchHighlight_f7a6a6e3{background-color:Highlight;color:HighlightText}}.categoryTabs_f7a6a6e3{-webkit-overflow-scrolling:touch;-ms-overflow-style:none;background-color:\"[theme:bodyBackground, default: #ffffff]\";background-color:var(--bodyBackground,#fff);border-bottom:1px solid;border-bottom-color:var(--neutralLight,#eaeaea);display:flex;flex-wrap:wrap;gap:4px;overflow-x:auto;padding:12px 20px;scrollbar-width:none}.categoryTabs_f7a6a6e3::-webkit-scrollbar{display:none}.categoryTab_f7a6a6e3{align-items:center;background:0 0;border:none;border-bottom:2px solid transparent;border-radius:4px 4px 0 0;color:\"[theme:bodyText, default: #323130]\";color:var(--bodyText,#323130);cursor:pointer;display:inline-flex;font-size:14px;font-weight:500;justify-content:center;padding:8px 16px;transition:all .2s ease;white-space:nowrap}.categoryTab_f7a6a6e3:hover{background-color:\"[theme:neutralLighter, default: #f4f4f4]\";background-color:var(--neutralLighter,#f4f4f4)}.categoryTab_f7a6a6e3:focus{outline:2px solid;outline-color:var(--focusBorder,#0078d4);outline-offset:-2px}.activeTab_f7a6a6e3{background-color:\"[theme:neutralLighterAlt, default: #f8f8f8]\";background-color:var(--neutralLighterAlt,#f8f8f8);border-bottom-color:\"[theme:themePrimary, default: #0078d4]\";border-bottom-color:var(--themePrimary,#0078d4);color:\"[theme:themePrimary, default: #0078d4]\";color:var(--themePrimary,#0078d4);font-weight:600}.faqList_f7a6a6e3{margin:0;padding:0}.faqItem_f7a6a6e3{border-bottom:1px solid;border-bottom-color:var(--bodyDivider,#edebe9)}.faqItem_f7a6a6e3:last-child{border-bottom:none}.faqQuestion_f7a6a6e3{align-items:center;background-color:transparent;cursor:pointer;display:flex;justify-content:space-between;padding:20px 28px;transition:background-color .2s ease}.faqQuestion_f7a6a6e3:hover{background-color:\"[theme:listItemBackgroundHovered, default: #f3f2f1]\";background-color:var(--listItemBackgroundHovered,#f3f2f1)}.faqQuestion_f7a6a6e3:focus{outline:2px solid;outline-color:var(--focusBorder,#0078d4);outline-offset:-2px}.faqQuestionText_f7a6a6e3{color:\"[theme:bodyText, default: #323130]\";color:var(--bodyText,#323130);flex:1;font-weight:600;padding-right:16px}.faqChevron_f7a6a6e3{align-items:center;color:\"[theme:bodySubtext, default: #605e5c]\";color:var(--bodySubtext,#605e5c);display:flex;flex-shrink:0;height:24px;justify-content:center;transition:transform .3s ease;width:24px}.faqChevron_f7a6a6e3 svg{height:16px;width:16px}.faqItem_f7a6a6e3.expanded_f7a6a6e3 .faqChevron_f7a6a6e3{transform:rotate(180deg)}.faqAnswer_f7a6a6e3{max-height:0;opacity:0;overflow:hidden;padding:0 28px;transition:max-height .3s ease,opacity .3s ease,padding .3s ease}.faqItem_f7a6a6e3.expanded_f7a6a6e3 .faqAnswer_f7a6a6e3{max-height:2000px;opacity:1;padding:0 28px 20px}.faqAnswerContent_f7a6a6e3{color:\"[theme:bodySubtext, default: #605e5c]\";color:var(--bodySubtext,#605e5c);line-height:1.6;margin-bottom:16px}.faqAnswerContent_f7a6a6e3 p{margin:0 0 12px}.faqAnswerContent_f7a6a6e3 p:last-child{margin-bottom:0}.faqAnswerContent_f7a6a6e3 ol,.faqAnswerContent_f7a6a6e3 ul{margin:0 0 12px;padding-left:24px}.faqAnswerContent_f7a6a6e3 a{color:\"[theme:link, default: #0078d4]\";color:var(--link,#0078d4);text-decoration:none}.faqAnswerContent_f7a6a6e3 a:hover{color:\"[theme:linkHovered, default: #004578]\";color:var(--linkHovered,#004578);text-decoration:underline}.attachments_f7a6a6e3{background-color:\"[theme:neutralLighter, default: #f4f4f4]\";background-color:var(--neutralLighter,#f4f4f4);border-radius:6px;margin-top:12px;padding:12px 16px}.attachmentsLabel_f7a6a6e3{align-items:center;color:\"[theme:neutralTertiary, default: #a6a6a6]\";color:var(--neutralTertiary,#a6a6a6);display:flex;font-size:.75rem;font-weight:600;gap:6px;letter-spacing:.5px;margin-bottom:8px;text-transform:uppercase}.attachmentsLabel_f7a6a6e3 svg{flex-shrink:0}.attachmentLink_f7a6a6e3{align-items:center;color:\"[theme:link, default: #0078d4]\";color:var(--link,#0078d4);display:flex;font-size:.875rem;gap:8px;padding:6px 0;text-decoration:none;transition:color .2s ease}.attachmentLink_f7a6a6e3:hover{color:\"[theme:linkHovered, default: #004578]\";color:var(--linkHovered,#004578);text-decoration:underline}.attachmentLink_f7a6a6e3:focus{outline:2px solid;outline-color:var(--focusBorder,#0078d4);outline-offset:2px}.attachmentIcon_f7a6a6e3{align-items:center;display:flex;flex-shrink:0;height:20px;justify-content:center;width:20px}.attachmentIcon_f7a6a6e3 svg{height:16px;width:16px}.emptyState_f7a6a6e3{align-items:center;color:\"[theme:bodySubtext, default: #605e5c]\";color:var(--bodySubtext,#605e5c);display:flex;flex-direction:column;justify-content:center;padding:48px 24px;text-align:center}.emptyState_f7a6a6e3 svg{margin-bottom:16px;opacity:.6}.emptyState_f7a6a6e3 p{font-size:.9375rem;line-height:1.5;margin:0;max-width:300px}.loading_f7a6a6e3{align-items:center;display:flex;justify-content:center;padding:48px 24px}.loading_f7a6a6e3:after{animation:spin_f7a6a6e3 1s linear infinite;border:3px solid;border-radius:50%;border-top:3px solid \"[theme:themeprimary, default: #0078d4]\";content:\"\";height:32px;width:32px}@keyframes spin_f7a6a6e3{to{transform:rotate(1turn)}}@media (forced-colors:active){.faqQuestion_f7a6a6e3{border:1px solid ButtonText}.faqItem_f7a6a6e3.expanded_f7a6a6e3 .faqQuestion_f7a6a6e3,.faqQuestion_f7a6a6e3:hover{background-color:Highlight;color:HighlightText}.attachmentLink_f7a6a6e3,.attachmentLink_f7a6a6e3:hover{color:LinkText}.categoryTab_f7a6a6e3{border:1px solid ButtonText}.activeTab_f7a6a6e3,.categoryTab_f7a6a6e3:hover{background-color:Highlight;color:HighlightText}.searchInput_f7a6a6e3{border:1px solid ButtonText}}.faqQuestion_f7a6a6e3:focus-visible{outline:2px solid;outline-color:var(--focusBorder,#0078d4);outline-offset:-2px}.attachmentLink_f7a6a6e3:focus-visible{outline:2px solid;outline-color:var(--focusBorder,#0078d4);outline-offset:2px}.categoryTab_f7a6a6e3:focus-visible{outline:2px solid;outline-color:var(--focusBorder,#0078d4);outline-offset:-2px}@media screen and (max-width:640px){.faqHeader_f7a6a6e3{padding:20px}.faqQuestion_f7a6a6e3{padding:16px 20px}.faqItem_f7a6a6e3.expanded_f7a6a6e3 .faqAnswer_f7a6a6e3{padding:0 20px 16px}.attachments_f7a6a6e3{padding:10px 12px}.emptyState_f7a6a6e3{padding:32px 20px}.categoryTabs_f7a6a6e3{padding:10px 16px}.categoryTab_f7a6a6e3{font-size:13px;padding:6px 12px}.searchSection_f7a6a6e3{padding:12px 16px}.searchInHeader_f7a6a6e3{margin-top:12px}.resultsInfo_f7a6a6e3{padding:6px 16px}}@media print{.customFaq_f7a6a6e3{border:1px solid #ccc;box-shadow:none}.faqHeader_f7a6a6e3{background:#333!important;-webkit-print-color-adjust:exact;print-color-adjust:exact}.faqAnswer_f7a6a6e3{max-height:none!important;opacity:1!important;padding:0 28px 20px!important}.faqChevron_f7a6a6e3{display:none}.faqQuestion_f7a6a6e3{cursor:default}.categoryTabs_f7a6a6e3,.resultsInfo_f7a6a6e3,.searchInHeader_f7a6a6e3,.searchSection_f7a6a6e3{display:none}.searchHighlight_f7a6a6e3{background-color:#ff0!important;-webkit-print-color-adjust:exact;print-color-adjust:exact}}\n/*# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJzb3VyY2VzIjpbImZpbGU6Ly8vVXNlcnMvdmlzaG51L0Rlc2t0b3AvY29kZWJhc2UvY3VzdG9tLWZhcS9zcmMvd2VicGFydHMvY3VzdG9tRmFxL0N1c3RvbUZhcS5tb2R1bGUuc2NzcyJdLCJuYW1lcyI6W10sIm1hcHBpbmdzIjoiQUFPQSxvQkFNRSwyREFBQSxDQUNBLDJDQUFBLENBTEEsaUJBQUEsQ0FFQSxvQ0FBQSxDQUhBLHNGQUFBLENBRUEsZUFJQSxDQU9GLG9CQUtFLG1EQUFBLENBQ0Esc0NBQUEsQ0FKQSx1Q0FBQSxDQUNBLHVCQUFBLENBRkEsaUJBS0EsQ0FHRixtQkFDRSxlQUFBLENBRUEsZUFBQSxDQURBLGNBQ0EsQ0FHRix5QkFHRSxlQUFBLENBRkEsUUFBQSxDQUNBLFVBQ0EsQ0FPRix3QkFLRSwyREFBQSxDQUNBLDhDQUFBLENBSEEsdUJBQUEsQ0FBQSw4Q0FBQSxDQUZBLGlCQUtBLENBR0YsMEJBRUUsY0FBQSxDQURBLGlCQUNBLENBR0YscUJBU0UsNkNBQUEsQ0FDQSxnQ0FBQSxDQUpBLFdBQUEsQ0FKQSxTQUFBLENBS0EsbUJBQUEsQ0FOQSxpQkFBQSxDQUVBLE9BQUEsQ0FDQSwwQkFBQSxDQUNBLFVBS0EsQ0FHRixzQkFVRSw0REFBQSxDQUNBLDRDQUFBLENBUkEsZ0JBQUEsQ0FDQSx1Q0FBQSxDQUNBLGlCQUFBLENBUUEsMENBQUEsQ0FDQSw2QkFBQSxDQVBBLG1CQUFBLENBREEsY0FBQSxDQUpBLGlCQUFBLENBTUEsMENBQUEsQ0FQQSxVQWFBLENBRUEsbUNBQ0Usc0RBQUEsQ0FDQSx5Q0FBQSxDQUZGLDRDQUNFLHNEQUFBLENBQ0EseUNBQUEsQ0FGRixtQ0FDRSxzREFBQSxDQUNBLHlDQUFBLENBR0YsNEJBRUUscURBQUEsQ0FDQSx3Q0FBQSxDQUNBLGdEQUFBLENBSEEsU0FHQSxDQUlKLHNCQVdFLGtCQUFBLENBTkEsY0FBQSxDQUNBLFdBQUEsQ0FFQSxpQkFBQSxDQU1BLDZDQUFBLENBQ0EsZ0NBQUEsQ0FOQSxjQUFBLENBQ0EsWUFBQSxDQUVB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 */", true);

// Exports
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = ({
  customFaq_f7a6a6e3: "customFaq_f7a6a6e3",
  faqHeader_f7a6a6e3: "faqHeader_f7a6a6e3",
  faqTitle_f7a6a6e3: "faqTitle_f7a6a6e3",
  faqDescription_f7a6a6e3: "faqDescription_f7a6a6e3",
  searchSection_f7a6a6e3: "searchSection_f7a6a6e3",
  searchContainer_f7a6a6e3: "searchContainer_f7a6a6e3",
  searchIcon_f7a6a6e3: "searchIcon_f7a6a6e3",
  searchInput_f7a6a6e3: "searchInput_f7a6a6e3",
  searchInput: "searchInput",
  clearButton_f7a6a6e3: "clearButton_f7a6a6e3",
  searchInHeader_f7a6a6e3: "searchInHeader_f7a6a6e3",
  searchInHeader: "searchInHeader",
  resultsInfo_f7a6a6e3: "resultsInfo_f7a6a6e3",
  searchHighlight_f7a6a6e3: "searchHighlight_f7a6a6e3",
  categoryTabs_f7a6a6e3: "categoryTabs_f7a6a6e3",
  categoryTab_f7a6a6e3: "categoryTab_f7a6a6e3",
  activeTab_f7a6a6e3: "activeTab_f7a6a6e3",
  faqList_f7a6a6e3: "faqList_f7a6a6e3",
  faqItem_f7a6a6e3: "faqItem_f7a6a6e3",
  faqQuestion_f7a6a6e3: "faqQuestion_f7a6a6e3",
  faqQuestionText_f7a6a6e3: "faqQuestionText_f7a6a6e3",
  faqChevron_f7a6a6e3: "faqChevron_f7a6a6e3",
  expanded_f7a6a6e3: "expanded_f7a6a6e3",
  faqAnswer_f7a6a6e3: "faqAnswer_f7a6a6e3",
  faqAnswerContent_f7a6a6e3: "faqAnswerContent_f7a6a6e3",
  attachments_f7a6a6e3: "attachments_f7a6a6e3",
  attachmentsLabel_f7a6a6e3: "attachmentsLabel_f7a6a6e3",
  attachmentLink_f7a6a6e3: "attachmentLink_f7a6a6e3",
  attachmentIcon_f7a6a6e3: "attachmentIcon_f7a6a6e3",
  emptyState_f7a6a6e3: "emptyState_f7a6a6e3",
  loading_f7a6a6e3: "loading_f7a6a6e3",
  spin_f7a6a6e3: "spin_f7a6a6e3"
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
    customFaq: 'customFaq_f7a6a6e3',
    faqHeader: 'faqHeader_f7a6a6e3',
    faqTitle: 'faqTitle_f7a6a6e3',
    faqDescription: 'faqDescription_f7a6a6e3',
    searchSection: 'searchSection_f7a6a6e3',
    searchContainer: 'searchContainer_f7a6a6e3',
    searchIcon: 'searchIcon_f7a6a6e3',
    searchInput: 'searchInput_f7a6a6e3',
    clearButton: 'clearButton_f7a6a6e3',
    searchInHeader: 'searchInHeader_f7a6a6e3',
    resultsInfo: 'resultsInfo_f7a6a6e3',
    searchHighlight: 'searchHighlight_f7a6a6e3',
    categoryTabs: 'categoryTabs_f7a6a6e3',
    categoryTab: 'categoryTab_f7a6a6e3',
    activeTab: 'activeTab_f7a6a6e3',
    faqList: 'faqList_f7a6a6e3',
    faqItem: 'faqItem_f7a6a6e3',
    faqQuestion: 'faqQuestion_f7a6a6e3',
    faqQuestionText: 'faqQuestionText_f7a6a6e3',
    faqChevron: 'faqChevron_f7a6a6e3',
    expanded: 'expanded_f7a6a6e3',
    faqAnswer: 'faqAnswer_f7a6a6e3',
    faqAnswerContent: 'faqAnswerContent_f7a6a6e3',
    attachments: 'attachments_f7a6a6e3',
    attachmentsLabel: 'attachmentsLabel_f7a6a6e3',
    attachmentLink: 'attachmentLink_f7a6a6e3',
    attachmentIcon: 'attachmentIcon_f7a6a6e3',
    emptyState: 'emptyState_f7a6a6e3',
    loading: 'loading_f7a6a6e3',
    spin: 'spin_f7a6a6e3'
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
            for (var i = 0; i < data.value.length; i++) {
                lists.push({
                    id: data.value[i].Id,
                    title: data.value[i].Title
                });
            }
            return lists;
        });
    };
    /**
     * Get columns for a specific list
     * Returns Text, Note, and Choice columns
     */
    SharePointService.prototype.getListColumns = function (listId) {
        // Filter for Text, Note, and Choice field types
        var endpoint = this._siteUrl + '/_api/web/lists(guid\'' + listId + '\')/fields?$filter=(TypeAsString eq \'Text\' or TypeAsString eq \'Note\' or TypeAsString eq \'Choice\') and Hidden eq false and ReadOnlyField eq false&$select=Id,InternalName,Title,TypeAsString&$orderby=Title';
        return this._context.spHttpClient.get(endpoint, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__.SPHttpClient.configurations.v1)
            .then(function (response) {
            if (!response.ok) {
                throw new Error('Error fetching columns: ' + response.statusText);
            }
            return response.json();
        })
            .then(function (data) {
            var columns = [];
            for (var i = 0; i < data.value.length; i++) {
                columns.push({
                    id: data.value[i].Id,
                    internalName: data.value[i].InternalName,
                    title: data.value[i].Title,
                    type: data.value[i].TypeAsString
                });
            }
            return columns;
        });
    };
    /**
     * Get all items from a list with specified columns
     * Also fetches attachments for each item
     */
    SharePointService.prototype.getListItems = function (listId, titleColumn, descriptionColumn, categoryColumn) {
        // Build select query - always include Id and Attachments
        var selectFields = ['Id', titleColumn];
        if (descriptionColumn !== titleColumn) {
            selectFields.push(descriptionColumn);
        }
        if (categoryColumn && categoryColumn !== titleColumn && categoryColumn !== descriptionColumn) {
            selectFields.push(categoryColumn);
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
            for (var i = 0; i < data.value.length; i++) {
                var item = data.value[i];
                itemPromises.push(self._processItem(item, listId, titleColumn, descriptionColumn, categoryColumn));
            }
            return Promise.all(itemPromises);
        })
            .then(function (items) {
            // Filter out items without titles
            var filteredItems = [];
            for (var i = 0; i < items.length; i++) {
                if (items[i].title && items[i].title.trim() !== '') {
                    filteredItems.push(items[i]);
                }
            }
            return filteredItems;
        });
    };
    /**
     * Process a single item and fetch its attachments
     */
    SharePointService.prototype._processItem = function (item, listId, titleColumn, descriptionColumn, categoryColumn) {
        var self = this;
        var faqItem = {
            id: item.Id,
            title: item[titleColumn] || '',
            description: item[descriptionColumn] || '',
            category: categoryColumn ? (item[categoryColumn] || '') : '',
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
            for (var i = 0; i < data.value.length; i++) {
                attachments.push({
                    fileName: data.value[i].FileName,
                    url: baseUrl + data.value[i].ServerRelativeUrl
                });
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
var _themeTokenRegex = /[\'\"]\[theme:\s*(\w+)\s*(?:\,\s*default:\s*([\\"\']?[\.\,\(\)\#\-\s\w]*[\.\,\(\)\#\-\w][\"\']?))?\s*\][\'\"]/g;
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
                for (var i = 0; i < semanticKeys.length; i++) {
                    var key = semanticKeys[i];
                    var value = semanticColors[key];
                    if (value) {
                        this.domElement.style.setProperty('--' + key, value);
                    }
                }
            }
            if (theme.palette) {
                var palette = theme.palette;
                var paletteKeys = Object.keys(palette);
                for (var i = 0; i < paletteKeys.length; i++) {
                    var key = paletteKeys[i];
                    var value = palette[key];
                    if (value) {
                        this.domElement.style.setProperty('--' + key, value);
                    }
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
        for (var i = 0; i < this._faqItems.length; i++) {
            var category = this._faqItems[i].category;
            if (category && !categorySet[category]) {
                categorySet[category] = true;
                this._categories.push(category);
            }
        }
        this._categories.sort();
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
        else {
            for (var i = 0; i < this._faqItems.length; i++) {
                if (this._faqItems[i].category === this._selectedCategory) {
                    filtered.push(this._faqItems[i]);
                }
            }
        }
        // Then filter by search query
        if (this._searchQuery && this._searchQuery.trim() !== '') {
            var query = this._searchQuery.toLowerCase().trim();
            var searchResults = [];
            for (var i = 0; i < filtered.length; i++) {
                var item = filtered[i];
                var titleMatch = item.title && item.title.toLowerCase().indexOf(query) !== -1;
                var answerMatch = false;
                if (this.properties.searchInAnswers && item.description) {
                    // Strip HTML tags for search
                    var plainDescription = item.description.replace(/<[^>]*>/g, '');
                    answerMatch = plainDescription.toLowerCase().indexOf(query) !== -1;
                }
                if (titleMatch || answerMatch) {
                    searchResults.push(item);
                }
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
        var count = 0;
        for (var i = 0; i < this._faqItems.length; i++) {
            if (this._faqItems[i].category === this._selectedCategory) {
                count++;
            }
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
            // For HTML content, we need to be careful not to highlight inside tags
            // Simple approach: highlight only in text nodes
            var regex_1 = new RegExp('(' + this._escapeRegExp(query) + ')', 'gi');
            return text.replace(/>([^<]+)</g, function (match, content) {
                return '>' + content.replace(regex_1, '<mark class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].searchHighlight + '">$1</mark>') + '<';
            });
        }
        else {
            var regex = new RegExp('(' + this._escapeRegExp(query) + ')', 'gi');
            return text.replace(regex, '<mark class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].searchHighlight + '">$1</mark>');
        }
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
                // Build category tabs
                if (this.properties.categoryColumn && this._categories.length > 0) {
                    var tabsArray = [];
                    var allActiveClass = this._selectedCategory === 'All' ? ' ' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].activeTab : '';
                    tabsArray.push('<button class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].categoryTab + allActiveClass + '" ' +
                        'data-category="All" ' +
                        'style="color: ' + (this._selectedCategory === 'All' ? headerBgColor : bodyText) + '; ' +
                        'border-bottom-color: ' + (this._selectedCategory === 'All' ? headerBgColor : 'transparent') + ';">' +
                        'All' +
                        '</button>');
                    for (var i = 0; i < this._categories.length; i++) {
                        var category = this._categories[i];
                        var isActive = this._selectedCategory === category;
                        var activeClass = isActive ? ' ' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].activeTab : '';
                        tabsArray.push('<button class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].categoryTab + activeClass + '" ' +
                            'data-category="' + this._escapeHtml(category) + '" ' +
                            'style="color: ' + (isActive ? headerBgColor : bodyText) + '; ' +
                            'border-bottom-color: ' + (isActive ? headerBgColor : 'transparent') + ';">' +
                            this._escapeHtml(category) +
                            '</button>');
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
                    for (var index = 0; index < filteredItems.length; index++) {
                        var item = filteredItems[index];
                        var attachmentsHtml = '';
                        if (item.attachments && item.attachments.length > 0) {
                            var attachmentLinksArray = [];
                            for (var j = 0; j < item.attachments.length; j++) {
                                var att = item.attachments[j];
                                attachmentLinksArray.push('<a href="' + att.url + '" target="_blank" rel="noopener noreferrer" class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].attachmentLink + '" style="color: ' + linkColor + ';">' +
                                    '<span class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].attachmentIcon + '">' +
                                    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
                                    '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>' +
                                    '<polyline points="14 2 14 8 20 8"/>' +
                                    '</svg>' +
                                    '</span>' +
                                    this._escapeHtml(att.fileName) +
                                    '</a>');
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
            // Handle Enter key
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
        var _loop_1 = function (i) {
            var tab = tabs[i];
            tab.addEventListener('click', function () {
                var category = tab.getAttribute('data-category');
                if (category) {
                    self._selectedCategory = category;
                    self.render();
                }
            });
        };
        for (var i = 0; i < tabs.length; i++) {
            _loop_1(i);
        }
    };
    CustomFaqWebPart.prototype._attachEventListeners = function () {
        var faqItems = this.domElement.querySelectorAll('.' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqItem);
        var questions = this.domElement.querySelectorAll('.' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqQuestion);
        var self = this;
        var _loop_2 = function (index) {
            var questionElement = questions[index];
            questionElement.addEventListener('mouseenter', function () {
                var hoverBg = questionElement.getAttribute('data-hover-bg');
                if (hoverBg) {
                    questionElement.style.backgroundColor = hoverBg;
                }
            });
            questionElement.addEventListener('mouseleave', function () {
                questionElement.style.backgroundColor = '';
            });
            (function (idx) {
                questionElement.addEventListener('click', function () {
                    var faqItem = faqItems[idx];
                    if (!self.properties.allowMultipleExpanded) {
                        for (var i = 0; i < faqItems.length; i++) {
                            if (i !== idx && faqItems[i].classList.contains(_CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].expanded)) {
                                faqItems[i].classList.remove(_CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].expanded);
                            }
                        }
                    }
                    if (faqItem.classList.contains(_CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].expanded)) {
                        faqItem.classList.remove(_CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].expanded);
                    }
                    else {
                        faqItem.classList.add(_CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].expanded);
                    }
                });
            })(index);
        };
        for (var index = 0; index < questions.length; index++) {
            _loop_2(index);
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
        if (description.indexOf('<') !== -1 && description.indexOf('>') !== -1) {
            return description;
        }
        return this._escapeHtml(description).replace(/\n/g, '<br>');
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
        return this._spService.getListItems(this.properties.selectedList, this.properties.titleColumn, this.properties.descriptionColumn, this.properties.categoryColumn || undefined)
            .then(function (items) {
            _this._faqItems = items;
            _this._extractCategories();
            var categoryExists = false;
            for (var i = 0; i < _this._categories.length; i++) {
                if (_this._categories[i] === _this._selectedCategory) {
                    categoryExists = true;
                    break;
                }
            }
            if (_this._selectedCategory !== 'All' && !categoryExists) {
                _this._selectedCategory = 'All';
            }
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
            this._faqItems = [];
            this._categories = [];
            this._selectedCategory = 'All';
            this._searchQuery = '';
            this._loadColumns().then(function () {
                self.context.propertyPane.refresh();
                self.render();
            });
        }
        else if ((propertyPath === 'titleColumn' || propertyPath === 'descriptionColumn' || propertyPath === 'categoryColumn') && newValue !== oldValue) {
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
        for (var i = 0; i < this._lists.length; i++) {
            options.push({ key: this._lists[i].id, text: this._lists[i].title });
        }
        return options;
    };
    CustomFaqWebPart.prototype._getTitleColumnOptions = function () {
        var options = [
            { key: '', text: '-- Select a column --' }
        ];
        for (var i = 0; i < this._columns.length; i++) {
            var col = this._columns[i];
            if (col.type === 'Text' || col.type === 'Note') {
                options.push({ key: col.internalName, text: col.title });
            }
        }
        return options;
    };
    CustomFaqWebPart.prototype._getDescriptionColumnOptions = function () {
        var options = [
            { key: '', text: '-- Select a column --' }
        ];
        for (var i = 0; i < this._columns.length; i++) {
            var col = this._columns[i];
            if (col.type === 'Text' || col.type === 'Note') {
                var typeLabel = col.type === 'Note' ? 'Multi-line' : 'Single-line';
                options.push({ key: col.internalName, text: col.title + ' (' + typeLabel + ')' });
            }
        }
        return options;
    };
    CustomFaqWebPart.prototype._getCategoryColumnOptions = function () {
        var options = [
            { key: '', text: '-- No category (disable tabs) --' }
        ];
        for (var i = 0; i < this._columns.length; i++) {
            var col = this._columns[i];
            if (col.type === 'Text' || col.type === 'Choice') {
                options.push({ key: col.internalName, text: col.title });
            }
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