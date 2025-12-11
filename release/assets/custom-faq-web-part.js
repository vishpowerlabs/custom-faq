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


_node_modules_microsoft_sp_css_loader_node_modules_microsoft_load_themed_styles_lib_es6_index_js__WEBPACK_IMPORTED_MODULE_0__.loadStyles(".customFaq_aaee6cf9{background-color:\"[theme:bodyBackground, default: #ffffff]\";background-color:var(--bodyBackground,#fff);border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,.08);font-family:Segoe UI,-apple-system,BlinkMacSystemFont,Roboto,Helvetica Neue,sans-serif;overflow:hidden}.faqHeader_aaee6cf9{background:\"[theme:themePrimary, default: #0078d4]\";background:var(--themePrimary,#0078d4);color:\"[theme:white, default: #ffffff]\";color:var(--white,#fff);padding:24px 28px}.faqTitle_aaee6cf9{font-weight:600;line-height:1.3;margin:0 0 4px}.faqDescription_aaee6cf9{line-height:1.4;margin:0;opacity:.9}.categoryTabs_aaee6cf9{-webkit-overflow-scrolling:touch;-ms-overflow-style:none;background-color:\"[theme:bodyBackground, default: #ffffff]\";background-color:var(--bodyBackground,#fff);border-bottom:1px solid;border-bottom-color:var(--neutralLight,#eaeaea);display:flex;flex-wrap:wrap;gap:4px;overflow-x:auto;padding:12px 20px;scrollbar-width:none}.categoryTabs_aaee6cf9::-webkit-scrollbar{display:none}.categoryTab_aaee6cf9{align-items:center;background:0 0;border:none;border-bottom:2px solid transparent;border-radius:4px 4px 0 0;color:\"[theme:bodyText, default: #323130]\";color:var(--bodyText,#323130);cursor:pointer;display:inline-flex;font-size:14px;font-weight:500;justify-content:center;padding:8px 16px;transition:all .2s ease;white-space:nowrap}.categoryTab_aaee6cf9:hover{background-color:\"[theme:neutralLighter, default: #f4f4f4]\";background-color:var(--neutralLighter,#f4f4f4)}.categoryTab_aaee6cf9:focus{outline:2px solid;outline-color:var(--focusBorder,#0078d4);outline-offset:-2px}.activeTab_aaee6cf9{background-color:\"[theme:neutralLighterAlt, default: #f8f8f8]\";background-color:var(--neutralLighterAlt,#f8f8f8);border-bottom-color:\"[theme:themePrimary, default: #0078d4]\";border-bottom-color:var(--themePrimary,#0078d4);color:\"[theme:themePrimary, default: #0078d4]\";color:var(--themePrimary,#0078d4);font-weight:600}.faqList_aaee6cf9{margin:0;padding:0}.faqItem_aaee6cf9{border-bottom:1px solid;border-bottom-color:var(--bodyDivider,#edebe9)}.faqItem_aaee6cf9:last-child{border-bottom:none}.faqQuestion_aaee6cf9{align-items:center;background-color:transparent;cursor:pointer;display:flex;justify-content:space-between;padding:20px 28px;transition:background-color .2s ease}.faqQuestion_aaee6cf9:hover{background-color:\"[theme:listItemBackgroundHovered, default: #f3f2f1]\";background-color:var(--listItemBackgroundHovered,#f3f2f1)}.faqQuestion_aaee6cf9:focus{outline:2px solid;outline-color:var(--focusBorder,#0078d4);outline-offset:-2px}.faqQuestionText_aaee6cf9{color:\"[theme:bodyText, default: #323130]\";color:var(--bodyText,#323130);flex:1;font-weight:600;padding-right:16px}.faqChevron_aaee6cf9{align-items:center;color:\"[theme:bodySubtext, default: #605e5c]\";color:var(--bodySubtext,#605e5c);display:flex;flex-shrink:0;height:24px;justify-content:center;transition:transform .3s ease;width:24px}.faqChevron_aaee6cf9 svg{height:16px;width:16px}.faqItem_aaee6cf9.expanded_aaee6cf9 .faqChevron_aaee6cf9{transform:rotate(180deg)}.faqAnswer_aaee6cf9{max-height:0;opacity:0;overflow:hidden;padding:0 28px;transition:max-height .3s ease,opacity .3s ease,padding .3s ease}.faqItem_aaee6cf9.expanded_aaee6cf9 .faqAnswer_aaee6cf9{max-height:2000px;opacity:1;padding:0 28px 20px}.faqAnswerContent_aaee6cf9{color:\"[theme:bodySubtext, default: #605e5c]\";color:var(--bodySubtext,#605e5c);line-height:1.6;margin-bottom:16px}.faqAnswerContent_aaee6cf9 p{margin:0 0 12px}.faqAnswerContent_aaee6cf9 p:last-child{margin-bottom:0}.faqAnswerContent_aaee6cf9 ol,.faqAnswerContent_aaee6cf9 ul{margin:0 0 12px;padding-left:24px}.faqAnswerContent_aaee6cf9 a{color:\"[theme:link, default: #0078d4]\";color:var(--link,#0078d4);text-decoration:none}.faqAnswerContent_aaee6cf9 a:hover{color:\"[theme:linkHovered, default: #004578]\";color:var(--linkHovered,#004578);text-decoration:underline}.attachments_aaee6cf9{background-color:\"[theme:neutralLighter, default: #f4f4f4]\";background-color:var(--neutralLighter,#f4f4f4);border-radius:6px;margin-top:12px;padding:12px 16px}.attachmentsLabel_aaee6cf9{align-items:center;color:\"[theme:neutralTertiary, default: #a6a6a6]\";color:var(--neutralTertiary,#a6a6a6);display:flex;font-size:.75rem;font-weight:600;gap:6px;letter-spacing:.5px;margin-bottom:8px;text-transform:uppercase}.attachmentsLabel_aaee6cf9 svg{flex-shrink:0}.attachmentLink_aaee6cf9{align-items:center;color:\"[theme:link, default: #0078d4]\";color:var(--link,#0078d4);display:flex;font-size:.875rem;gap:8px;padding:6px 0;text-decoration:none;transition:color .2s ease}.attachmentLink_aaee6cf9:hover{color:\"[theme:linkHovered, default: #004578]\";color:var(--linkHovered,#004578);text-decoration:underline}.attachmentLink_aaee6cf9:focus{outline:2px solid;outline-color:var(--focusBorder,#0078d4);outline-offset:2px}.attachmentIcon_aaee6cf9{align-items:center;display:flex;flex-shrink:0;height:20px;justify-content:center;width:20px}.attachmentIcon_aaee6cf9 svg{height:16px;width:16px}.emptyState_aaee6cf9{align-items:center;color:\"[theme:bodySubtext, default: #605e5c]\";color:var(--bodySubtext,#605e5c);display:flex;flex-direction:column;justify-content:center;padding:48px 24px;text-align:center}.emptyState_aaee6cf9 svg{margin-bottom:16px;opacity:.6}.emptyState_aaee6cf9 p{font-size:.9375rem;line-height:1.5;margin:0;max-width:300px}.loading_aaee6cf9{align-items:center;display:flex;justify-content:center;padding:48px 24px}.loading_aaee6cf9:after{animation:spin_aaee6cf9 1s linear infinite;border:3px solid;border-radius:50%;border-top:3px solid \"[theme:themeprimary, default: #0078d4]\";content:\"\";height:32px;width:32px}@keyframes spin_aaee6cf9{to{transform:rotate(1turn)}}@media (forced-colors:active){.faqQuestion_aaee6cf9{border:1px solid ButtonText}.faqItem_aaee6cf9.expanded_aaee6cf9 .faqQuestion_aaee6cf9,.faqQuestion_aaee6cf9:hover{background-color:Highlight;color:HighlightText}.attachmentLink_aaee6cf9,.attachmentLink_aaee6cf9:hover{color:LinkText}.categoryTab_aaee6cf9{border:1px solid ButtonText}.activeTab_aaee6cf9,.categoryTab_aaee6cf9:hover{background-color:Highlight;color:HighlightText}}.faqQuestion_aaee6cf9:focus-visible{outline:2px solid;outline-color:var(--focusBorder,#0078d4);outline-offset:-2px}.attachmentLink_aaee6cf9:focus-visible{outline:2px solid;outline-color:var(--focusBorder,#0078d4);outline-offset:2px}.categoryTab_aaee6cf9:focus-visible{outline:2px solid;outline-color:var(--focusBorder,#0078d4);outline-offset:-2px}@media screen and (max-width:640px){.faqHeader_aaee6cf9{padding:20px}.faqQuestion_aaee6cf9{padding:16px 20px}.faqItem_aaee6cf9.expanded_aaee6cf9 .faqAnswer_aaee6cf9{padding:0 20px 16px}.attachments_aaee6cf9{padding:10px 12px}.emptyState_aaee6cf9{padding:32px 20px}.categoryTabs_aaee6cf9{padding:10px 16px}.categoryTab_aaee6cf9{font-size:13px;padding:6px 12px}}@media print{.customFaq_aaee6cf9{border:1px solid #ccc;box-shadow:none}.faqHeader_aaee6cf9{background:#333!important;-webkit-print-color-adjust:exact;print-color-adjust:exact}.faqAnswer_aaee6cf9{max-height:none!important;opacity:1!important;padding:0 28px 20px!important}.faqChevron_aaee6cf9{display:none}.faqQuestion_aaee6cf9{cursor:default}.categoryTabs_aaee6cf9{display:none}}\n/*# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJzb3VyY2VzIjpbImZpbGU6Ly8vVXNlcnMvdmlzaG51L0Rlc2t0b3AvY29kZWJhc2UvY3VzdG9tLWZhcS9zcmMvd2VicGFydHMvY3VzdG9tRmFxL0N1c3RvbUZhcS5tb2R1bGUuc2NzcyJdLCJuYW1lcyI6W10sIm1hcHBpbmdzIjoiQUFPQSxvQkFPRSwyREFBQSxDQUNBLDJDQUFBLENBTkEsaUJBQUEsQ0FFQSxvQ0FBQSxDQUhBLHNGQUFBLENBRUEsZUFLQSxDQU9GLG9CQU1FLG1EQUFBLENBQ0Esc0NBQUEsQ0FMQSx1Q0FBQSxDQUNBLHVCQUFBLENBRkEsaUJBTUEsQ0FHRixtQkFDRSxlQUFBLENBRUEsZUFBQSxDQURBLGNBQ0EsQ0FHRix5QkFHRSxlQUFBLENBRkEsUUFBQSxDQUNBLFVBQ0EsQ0FPRix1QkFVRSxnQ0FBQSxDQUlBLHVCQUFBLENBUEEsMkRBQUEsQ0FDQSwyQ0FBQSxDQUZBLHVCQUFBLENBQUEsK0NBQUEsQ0FMQSxZQUFBLENBQ0EsY0FBQSxDQUNBLE9BQUEsQ0FNQSxlQUFBLENBTEEsaUJBQUEsQ0FTQSxvQkFDQSxDQUVBLDBDQUNFLFlBQUEsQ0FJSixzQkFFRSxrQkFBQSxDQUtBLGNBQUEsQ0FEQSxXQUFBLENBQUEsbUNBQUEsQ0FPQSx5QkFBQSxDQUVBLDBDQUFBLENBQ0EsNkJBQUEsQ0FSQSxjQUFBLENBUEEsbUJBQUEsQ0FRQSxjQUFBLENBQ0EsZUFBQSxDQVBBLHNCQUFBLENBQ0EsZ0JBQUEsQ0FRQSx1QkFBQSxDQURBLGtCQUtBLENBRUEsNEJBQ0UsMkRBQUEsQ0FDQSw4Q0FBQSxDQUdGLDRCQUNFLGlCQUFBLENBQ0Esd0NBQUEsQ0FDQSxtQkFBQSxDQUlKLG9CQVNFLDhEQUFBLENBQ0EsaURBQUEsQ0FKQSw0REFBQSxDQUNBLCtDQUFBLENBSkEsOENBQUEsQ0FDQSxpQ0FBQSxDQUhBLGVBU0EsQ0FPRixrQkFFRSxRQUFBLENBREEsU0FDQSxDQU9GLGtCQUVFLHVCQUFBLENBQUEsOENBQUEsQ0FFQSw2QkFDRSxrQkFBQSxDQVFKLHNCQUVFLGtCQUFBLENBTUEsNEJBQUEsQ0FIQSxjQUFBLENBSkEsWUFBQSxDQUVBLDZCQUFBLENBQ0EsaUJBQUEsQ0FFQSxvQ0FFQSxDQUVBLDRCQUNFLHNFQUFBLENBQ0EseURBQUEsQ0FHRiw0QkFDRSxpQkFBQSxDQUNBLHdDQUFBLENBQ0EsbUJBQUEsQ0FJSiwwQkFLRSwwQ0FBQSxDQUNBLDZCQUFBLENBSkEsTUFBQSxDQURBLGVBQUEsQ0FFQSxrQkFHQSxDQU9GLHFCQUlFLGtCQUFBLENBS0EsNkNBQUEsQ0FDQSxnQ0FBQSxDQVBBLFlBQUEsQ0FHQSxhQUFBLENBSkEsV0FBQSxDQUdBLHNCQUFBLENBRUEsNkJBQUEsQ0FOQSxVQVNBLENBRUEseUJBRUUsV0FBQSxDQURBLFVBQ0EsQ0FLSix5REFDRSx3QkFBQSxDQU9GLG9CQUNFLFlBQUEsQ0FHQSxTQUFBLENBRkEsZUFBQSxDQUNBLGNBQUEsQ0FFQSxnRUFBQSxDQUlGLHdEQUNFLGlCQUFBLENBQ0EsU0FBQSxDQUNBLG1CQUFBLENBR0YsMkJBSUUsNkNBQUEsQ0FDQSxnQ0FBQSxDQUpBLGVBQUEsQ0FDQSxrQkFHQSxDQUdBLDZCQUNFLGVBQUEsQ0FFQSx3Q0FDRSxlQUFBLENBSUosNERBQ0UsZUFBQSxDQUNBLGlCQUFBLENBR0YsNkJBQ0Usc0NBQUEsQ0FDQSx5QkFBQSxDQUNBLG9CQUFBLENBRUEsbUNBQ0UsNkNBQUEsQ0FDQSxnQ0FBQSxDQUNBLHlCQUFBLENBU04sc0JBS0UsMkRBQUEsQ0FDQSw4Q0FBQSxDQUxBLGlCQUFBLENBRUEsZUFBQSxDQURBLGlCQUlBLENBR0YsMkJBT0Usa0JBQUEsQ0FHQSxpREFBQSxDQUNBLG9DQUFBLENBTEEsWUFBQSxDQUxBLGdCQUFBLENBQ0EsZUFBQSxDQU1BLE9BQUEsQ0FKQSxtQkFBQSxDQUNBLGlCQUFBLENBRkEsd0JBUUEsQ0FFQSwrQkFDRSxhQUFBLENBSUoseUJBRUUsa0JBQUEsQ0FPQSxzQ0FBQSxDQUNBLHlCQUFBLENBVEEsWUFBQSxDQUdBLGlCQUFBLENBREEsT0FBQSxDQUdBLGFBQUEsQ0FEQSxvQkFBQSxDQUVBLHlCQUdBLENBRUEsK0JBQ0UsNkNBQUEsQ0FDQSxnQ0FBQSxDQUNBLHlCQUFBLENBR0YsK0JBQ0UsaUJBQUEsQ0FDQSx3Q0FBQSxDQUNBLGtCQUFBLENBSUoseUJBSUUsa0JBQUEsQ0FEQSxZQUFBLENBR0EsYUFBQSxDQUpBLFdBQUEsQ0FHQSxzQkFBQSxDQUpBLFVBS0EsQ0FFQSw2QkFFRSxXQUFBLENBREEsVUFDQSxDQVFKLHFCQUdFLGtCQUFBLENBS0EsNkNBQUEsQ0FDQSxnQ0FBQSxDQVJBLFlBQUEsQ0FDQSxxQkFBQSxDQUVBLHNCQUFBLENBQ0EsaUJBQUEsQ0FDQSxpQkFHQSxDQUVBLHlCQUNFLGtCQUFBLENBQ0EsVUFBQSxDQUdGLHVCQUVFLGtCQUFBLENBRUEsZUFBQSxDQUhBLFFBQUEsQ0FFQSxlQUNBLENBUUosa0JBRUUsa0JBQUEsQ0FEQSxZQUFBLENBRUEsc0JBQUEsQ0FDQSxpQkFBQSxDQUVBLHdCQU9FLDBDQUFBLENBRkEsZ0JBQUEsQ0FDQSxpQkFBQSxDQURBLDZEQUFBLENBSkEsVUFBQSxDQUVBLFdBQUEsQ0FEQSxVQUtBLENBSUoseUJBQ0UsR0FDRSx1QkFBQSxDQUFBLENBU0osOEJBQ0Usc0JBQ0UsMkJBQUEsQ0FRRixzRkFDRSwwQkFBQSxDQUNBLG1CQUFBLENBTUEsd0RBQ0UsY0FBQSxDQUlKLHNCQUNFLDJCQUFBLENBUUYsZ0RBQ0UsMEJBQUEsQ0FDQSxtQkFBQSxDQUFBLENBS0osb0NBQ0UsaUJBQUEsQ0FDQSx3Q0FBQSxDQUNBLG1CQUFBLENBR0YsdUNBQ0UsaUJBQUEsQ0FDQSx3Q0FBQSxDQUNBLGtCQUFBLENBR0Ysb0NBQ0UsaUJBQUEsQ0FDQSx3Q0FBQSxDQUNBLG1CQUFBLENBT0Ysb0NBQ0Usb0JBQ0UsWUFBQSxDQUdGLHNCQUNFLGlCQUFBLENBR0Ysd0RBQ0UsbUJBQUEsQ0FHRixzQkFDRSxpQkFBQSxDQUdGLHFCQUNFLGlCQUFBLENBR0YsdUJBQ0UsaUJBQUEsQ0FHRixzQkFFRSxjQUFBLENBREEsZ0JBQ0EsQ0FBQSxDQVFKLGFBQ0Usb0JBRUUscUJBQUEsQ0FEQSxlQUNBLENBR0Ysb0JBQ0UseUJBQUEsQ0FDQSxnQ0FBQSxDQUNBLHdCQUFBLENBR0Ysb0JBQ0UseUJBQUEsQ0FDQSxtQkFBQSxDQUNBLDZCQUFBLENBR0YscUJBQ0UsWUFBQSxDQUdGLHNCQUNFLGNBQUEsQ0FHRix1QkFDRSxZQUFBLENBQUEiLCJmaWxlIjoiQ3VzdG9tRmFxLm1vZHVsZS5jc3MifQ== */", true);

// Exports
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = ({
  customFaq_aaee6cf9: "customFaq_aaee6cf9",
  faqHeader_aaee6cf9: "faqHeader_aaee6cf9",
  faqTitle_aaee6cf9: "faqTitle_aaee6cf9",
  faqDescription_aaee6cf9: "faqDescription_aaee6cf9",
  categoryTabs_aaee6cf9: "categoryTabs_aaee6cf9",
  categoryTab_aaee6cf9: "categoryTab_aaee6cf9",
  activeTab_aaee6cf9: "activeTab_aaee6cf9",
  faqList_aaee6cf9: "faqList_aaee6cf9",
  faqItem_aaee6cf9: "faqItem_aaee6cf9",
  faqQuestion_aaee6cf9: "faqQuestion_aaee6cf9",
  faqQuestionText_aaee6cf9: "faqQuestionText_aaee6cf9",
  faqChevron_aaee6cf9: "faqChevron_aaee6cf9",
  expanded_aaee6cf9: "expanded_aaee6cf9",
  faqAnswer_aaee6cf9: "faqAnswer_aaee6cf9",
  faqAnswerContent_aaee6cf9: "faqAnswerContent_aaee6cf9",
  attachments_aaee6cf9: "attachments_aaee6cf9",
  attachmentsLabel_aaee6cf9: "attachmentsLabel_aaee6cf9",
  attachmentLink_aaee6cf9: "attachmentLink_aaee6cf9",
  attachmentIcon_aaee6cf9: "attachmentIcon_aaee6cf9",
  emptyState_aaee6cf9: "emptyState_aaee6cf9",
  loading_aaee6cf9: "loading_aaee6cf9",
  spin_aaee6cf9: "spin_aaee6cf9"
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
    customFaq: 'customFaq_aaee6cf9',
    faqHeader: 'faqHeader_aaee6cf9',
    faqTitle: 'faqTitle_aaee6cf9',
    faqDescription: 'faqDescription_aaee6cf9',
    categoryTabs: 'categoryTabs_aaee6cf9',
    categoryTab: 'categoryTab_aaee6cf9',
    activeTab: 'activeTab_aaee6cf9',
    faqList: 'faqList_aaee6cf9',
    faqItem: 'faqItem_aaee6cf9',
    faqQuestion: 'faqQuestion_aaee6cf9',
    faqQuestionText: 'faqQuestionText_aaee6cf9',
    faqChevron: 'faqChevron_aaee6cf9',
    expanded: 'expanded_aaee6cf9',
    faqAnswer: 'faqAnswer_aaee6cf9',
    faqAnswerContent: 'faqAnswerContent_aaee6cf9',
    attachments: 'attachments_aaee6cf9',
    attachmentsLabel: 'attachmentsLabel_aaee6cf9',
    attachmentLink: 'attachmentLink_aaee6cf9',
    attachmentIcon: 'attachmentIcon_aaee6cf9',
    emptyState: 'emptyState_aaee6cf9',
    loading: 'loading_aaee6cf9',
    spin: 'spin_aaee6cf9'
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
        return _this;
    }
    /**
     * Initialize the web part
     */
    CustomFaqWebPart.prototype.onInit = function () {
        var _this = this;
        // Initialize SharePoint service
        this._spService = new _services_SharePointService__WEBPACK_IMPORTED_MODULE_5__.SharePointService(this.context);
        // Consume the ThemeProvider service for section background support
        this._themeProvider = this.context.serviceScope.consume(_microsoft_sp_component_base__WEBPACK_IMPORTED_MODULE_3__.ThemeProvider.serviceKey);
        // Get the current theme variant
        this._themeVariant = this._themeProvider.tryGetTheme();
        // Apply CSS variables from theme
        if (this._themeVariant) {
            this._setCSSVariables(this._themeVariant);
        }
        // Register handler for theme changes
        this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);
        // Load initial data
        return this._loadLists().then(function () {
            if (_this.properties.selectedList) {
                return _this._loadColumns().then(function () {
                    return _this._loadFaqItems();
                });
            }
            return Promise.resolve();
        });
    };
    /**
     * Handle theme change events (section background changes)
     */
    CustomFaqWebPart.prototype._handleThemeChangedEvent = function (args) {
        this._themeVariant = args.theme;
        if (this._themeVariant) {
            this._setCSSVariables(this._themeVariant);
        }
        this.render();
    };
    /**
     * Convert theme semantic colors to CSS variables for use in SCSS
     */
    CustomFaqWebPart.prototype._setCSSVariables = function (theme) {
        if (!this.domElement) {
            return;
        }
        // Set semantic colors as CSS variables
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
        // Set palette colors as CSS variables
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
    };
    /**
     * Extract unique categories from FAQ items
     */
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
        // Sort categories alphabetically
        this._categories.sort();
    };
    /**
     * Get filtered FAQ items based on selected category
     */
    CustomFaqWebPart.prototype._getFilteredItems = function () {
        if (this._selectedCategory === 'All') {
            return this._faqItems;
        }
        var filtered = [];
        for (var i = 0; i < this._faqItems.length; i++) {
            if (this._faqItems[i].category === this._selectedCategory) {
                filtered.push(this._faqItems[i]);
            }
        }
        return filtered;
    };
    /**
     * Get font size value with px suffix
     */
    CustomFaqWebPart.prototype._getFontSize = function (size, defaultSize) {
        return (size || defaultSize) + 'px';
    };
    /**
     * Render the web part
     */
    CustomFaqWebPart.prototype.render = function () {
        // Build inline styles from theme for elements that need direct styling
        var semanticColors = this._themeVariant ? this._themeVariant.semanticColors : undefined;
        var palette = this._themeVariant ? this._themeVariant.palette : undefined;
        // Header background using theme primary color
        var headerBgColor = (palette && palette.themePrimary) ? palette.themePrimary : '#0078d4';
        var headerBgColorDark = (palette && palette.themeDark) ? palette.themeDark : '#005a9e';
        // Body colors
        var bodyBackground = (semanticColors && semanticColors.bodyBackground) ? semanticColors.bodyBackground : '#ffffff';
        var bodyText = (semanticColors && semanticColors.bodyText) ? semanticColors.bodyText : '#323130';
        var bodySubtext = (semanticColors && semanticColors.bodySubtext) ? semanticColors.bodySubtext : '#605e5c';
        // Interactive colors
        var linkColor = (semanticColors && semanticColors.link) ? semanticColors.link : ((palette && palette.themePrimary) ? palette.themePrimary : '#0078d4');
        // Border and divider colors
        var bodyDivider = (semanticColors && semanticColors.bodyDivider) ? semanticColors.bodyDivider : '#edebe9';
        // Hover states
        var listItemBackgroundHovered = (semanticColors && semanticColors.listItemBackgroundHovered) ? semanticColors.listItemBackgroundHovered : '#f3f2f1';
        // Card/container background
        var cardBackground = (semanticColors && semanticColors.cardStandoutBackground) ? semanticColors.cardStandoutBackground : bodyBackground;
        // Attachment area background
        var neutralLighter = (palette && palette.neutralLighter) ? palette.neutralLighter : '#f4f4f4';
        var neutralTertiary = (palette && palette.neutralTertiary) ? palette.neutralTertiary : '#a6a6a6';
        var neutralLight = (palette && palette.neutralLight) ? palette.neutralLight : '#eaeaea';
        // Font sizes
        var webpartTitleFontSize = this._getFontSize(this.properties.webpartTitleFontSize, '24');
        var webpartDescriptionFontSize = this._getFontSize(this.properties.webpartDescriptionFontSize, '14');
        var faqTitleFontSize = this._getFontSize(this.properties.faqTitleFontSize, '16');
        var faqDescriptionFontSize = this._getFontSize(this.properties.faqDescriptionFontSize, '14');
        var faqItemsHtml = '';
        var categoryTabsHtml = '';
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
            // Build category tabs if category column is selected
            if (this.properties.categoryColumn && this._categories.length > 0) {
                var tabsArray = [];
                // Add "All" tab
                var allActiveClass = this._selectedCategory === 'All' ? ' ' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].activeTab : '';
                tabsArray.push('<button class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].categoryTab + allActiveClass + '" ' +
                    'data-category="All" ' +
                    'style="color: ' + (this._selectedCategory === 'All' ? headerBgColor : bodyText) + '; ' +
                    'border-bottom-color: ' + (this._selectedCategory === 'All' ? headerBgColor : 'transparent') + ';">' +
                    'All' +
                    '</button>');
                // Add category tabs
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
            if (filteredItems.length === 0) {
                faqItemsHtml = '<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].emptyState + '" style="color: ' + bodySubtext + ';">' +
                    '<svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="' + neutralTertiary + '" stroke-width="1.5">' +
                    '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>' +
                    '<polyline points="14 2 14 8 20 8"/>' +
                    '<line x1="12" y1="18" x2="12" y2="12"/>' +
                    '<line x1="9" y1="15" x2="15" y2="15"/>' +
                    '</svg>' +
                    '<p>No FAQ items found' + (this._selectedCategory !== 'All' ? ' in this category' : '') + '.</p>' +
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
                    itemsHtmlArray.push('<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqItem + '" data-index="' + index + '">' +
                        '<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqQuestion + '" style="border-bottom-color: ' + bodyDivider + ';" data-hover-bg="' + listItemBackgroundHovered + '">' +
                        '<span class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqQuestionText + '" style="color: ' + bodyText + '; font-size: ' + faqTitleFontSize + ';">' +
                        this._escapeHtml(item.title) +
                        '</span>' +
                        '<span class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqChevron + '" style="color: ' + bodySubtext + ';">' +
                        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
                        '<path d="M6 9l6 6 6-6"/>' +
                        '</svg>' +
                        '</span>' +
                        '</div>' +
                        '<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqAnswer + '">' +
                        '<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqAnswerContent + '" style="color: ' + bodySubtext + '; font-size: ' + faqDescriptionFontSize + ';">' +
                        this._formatDescription(item.description) +
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
        if (this.properties.webpartTitle || this.properties.webpartDescription) {
            headerHtml = '<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqHeader + '" style="background: linear-gradient(135deg, ' + headerBgColor + ' 0%, ' + headerBgColorDark + ' 100%);">';
            if (this.properties.webpartTitle) {
                headerHtml += '<h2 class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqTitle + '" style="font-size: ' + webpartTitleFontSize + ';">' + this._escapeHtml(this.properties.webpartTitle) + '</h2>';
            }
            if (this.properties.webpartDescription) {
                headerHtml += '<p class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqDescription + '" style="font-size: ' + webpartDescriptionFontSize + ';">' + this._escapeHtml(this.properties.webpartDescription) + '</p>';
            }
            headerHtml += '</div>';
        }
        this.domElement.innerHTML = '<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].customFaq + '" style="background-color: ' + cardBackground + ';">' +
            headerHtml +
            categoryTabsHtml +
            '<div class="' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqList + '">' +
            faqItemsHtml +
            '</div>' +
            '</div>';
        // Attach event listeners for accordion functionality
        this._attachEventListeners();
        // Attach event listeners for category tabs
        this._attachTabEventListeners();
    };
    /**
     * Attach click event listeners to category tabs
     */
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
    /**
     * Attach click event listeners to FAQ items
     */
    CustomFaqWebPart.prototype._attachEventListeners = function () {
        var faqItems = this.domElement.querySelectorAll('.' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqItem);
        var questions = this.domElement.querySelectorAll('.' + _CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].faqQuestion);
        var self = this;
        var _loop_2 = function (index) {
            var questionElement = questions[index];
            // Add hover effect
            questionElement.addEventListener('mouseenter', function () {
                var hoverBg = questionElement.getAttribute('data-hover-bg');
                if (hoverBg) {
                    questionElement.style.backgroundColor = hoverBg;
                }
            });
            questionElement.addEventListener('mouseleave', function () {
                questionElement.style.backgroundColor = '';
            });
            // Add click handler for accordion - use closure to capture index
            (function (idx) {
                questionElement.addEventListener('click', function () {
                    var faqItem = faqItems[idx];
                    if (!self.properties.allowMultipleExpanded) {
                        // Close all other items
                        for (var i = 0; i < faqItems.length; i++) {
                            if (i !== idx && faqItems[i].classList.contains(_CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].expanded)) {
                                faqItems[i].classList.remove(_CustomFaq_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].expanded);
                            }
                        }
                    }
                    // Toggle current item
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
    /**
     * Escape HTML to prevent XSS
     */
    CustomFaqWebPart.prototype._escapeHtml = function (text) {
        if (!text) {
            return '';
        }
        var div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    };
    /**
     * Format description text (handle rich text or plain text)
     */
    CustomFaqWebPart.prototype._formatDescription = function (description) {
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
    };
    /**
     * Load all lists from the current site
     */
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
    /**
     * Load columns for the selected list
     */
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
    /**
     * Load FAQ items from the selected list
     */
    CustomFaqWebPart.prototype._loadFaqItems = function () {
        var _this = this;
        if (!this.properties.selectedList || !this.properties.titleColumn || !this.properties.descriptionColumn) {
            this._faqItems = [];
            return Promise.resolve();
        }
        return this._spService.getListItems(this.properties.selectedList, this.properties.titleColumn, this.properties.descriptionColumn, this.properties.categoryColumn)
            .then(function (items) {
            _this._faqItems = items;
            _this._extractCategories();
            // Reset to All if current category no longer exists
            if (_this._selectedCategory !== 'All' && _this._categories.indexOf(_this._selectedCategory) === -1) {
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
        /**
         * Get data version
         */
        get: function () {
            return _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_0__.Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    /**
     * Handle property pane field changes
     */
    CustomFaqWebPart.prototype.onPropertyPaneFieldChanged = function (propertyPath, oldValue, newValue) {
        var self = this;
        if (propertyPath === 'selectedList' && newValue !== oldValue) {
            // Reset column selections when list changes
            this.properties.titleColumn = '';
            this.properties.descriptionColumn = '';
            this.properties.categoryColumn = '';
            this._faqItems = [];
            this._categories = [];
            this._selectedCategory = 'All';
            // Load columns for new list
            this._loadColumns().then(function () {
                // Refresh property pane to show new columns
                self.context.propertyPane.refresh();
                self.render();
            });
        }
        else if ((propertyPath === 'titleColumn' || propertyPath === 'descriptionColumn' || propertyPath === 'categoryColumn') && newValue !== oldValue) {
            // Reload FAQ items when columns change
            this._loadFaqItems().then(function () {
                self.render();
            });
        }
        else {
            this.render();
        }
    };
    /**
     * Get dropdown options for lists
     */
    CustomFaqWebPart.prototype._getListOptions = function () {
        var options = [
            { key: '', text: '-- Select a list --' }
        ];
        for (var i = 0; i < this._lists.length; i++) {
            options.push({ key: this._lists[i].id, text: this._lists[i].title });
        }
        return options;
    };
    /**
     * Get dropdown options for title columns (text columns only)
     */
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
    /**
     * Get dropdown options for description columns (text and multi-line text columns)
     */
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
    /**
     * Get dropdown options for category column (text and choice columns)
     */
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
    /**
     * Get font size options
     */
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
    /**
     * Configure the property pane
     */
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