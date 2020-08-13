(window["wpJsonpOfficeFluidContainer"] = window["wpJsonpOfficeFluidContainer"] || []).push([["fluidTable"],{

/***/ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/DelayedRender.js":
/*!********************************************************************************************************************************************************************************************************************************!*\
  !*** C:/Users/riarya/Desktop/www/office-bohemia/common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/DelayedRender.js ***!
  \********************************************************************************************************************************************************************************************************************************/
/*! exports provided: DelayedRender */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "DelayedRender", function() { return DelayedRender; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/tslib/1.13.0/node_modules/tslib/tslib.es6.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_1___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_1__);


/**
 * Utility component for delaying the render of a child component after a given delay. This component
 * requires a single child component; don't pass in many components. Wrap multiple components in a DIV
 * if necessary.
 *
 * @public
 * {@docCategory DelayedRender}
 */
var DelayedRender = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(DelayedRender, _super);
    function DelayedRender(props) {
        var _this = _super.call(this, props) || this;
        _this.state = {
            isRendered: false,
        };
        return _this;
    }
    DelayedRender.prototype.componentDidMount = function () {
        var _this = this;
        var delay = this.props.delay;
        this._timeoutId = window.setTimeout(function () {
            _this.setState({
                isRendered: true,
            });
        }, delay);
    };
    DelayedRender.prototype.componentWillUnmount = function () {
        if (this._timeoutId) {
            clearTimeout(this._timeoutId);
        }
    };
    DelayedRender.prototype.render = function () {
        return this.state.isRendered ? react__WEBPACK_IMPORTED_MODULE_1__["Children"].only(this.props.children) : null;
    };
    DelayedRender.defaultProps = {
        delay: 0,
    };
    return DelayedRender;
}(react__WEBPACK_IMPORTED_MODULE_1__["Component"]));



/***/ }),

/***/ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/office-ui-fabric-react/7.121.6_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/office-ui-fabric-react/lib/components/Checkbox/Checkbox.base.js":
/*!***********************************************************************************************************************************************************************************************************************************************************!*\
  !*** C:/Users/riarya/Desktop/www/office-bohemia/common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/office-ui-fabric-react/7.121.6_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/office-ui-fabric-react/lib/components/Checkbox/Checkbox.base.js ***!
  \***********************************************************************************************************************************************************************************************************************************************************/
/*! exports provided: CheckboxBase */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "CheckboxBase", function() { return CheckboxBase; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/tslib/1.13.0/node_modules/tslib/tslib.es6.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_1___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_1__);
/* harmony import */ var _Utilities__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../../Utilities */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/classNamesFunction.js");
/* harmony import */ var _Utilities__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../../Utilities */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/useFocusRects.js");
/* harmony import */ var _Utilities__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../../Utilities */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/aria.js");
/* harmony import */ var _Utilities__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../../Utilities */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/initializeComponentRef.js");
/* harmony import */ var _Utilities__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../../Utilities */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/warn/warnMutuallyExclusive.js");
/* harmony import */ var _Utilities__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../../Utilities */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/getId.js");
/* harmony import */ var _Icon__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../../Icon */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/office-ui-fabric-react/7.121.6_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/office-ui-fabric-react/lib/components/Icon/Icon.js");
/* harmony import */ var _KeytipData__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ../../KeytipData */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/office-ui-fabric-react/7.121.6_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/office-ui-fabric-react/lib/components/KeytipData/KeytipData.js");





var getClassNames = Object(_Utilities__WEBPACK_IMPORTED_MODULE_2__["classNamesFunction"])();
var CheckboxBase = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(CheckboxBase, _super);
    /**
     * Initialize a new instance of the Checkbox
     * @param props - Props for the component
     * @param context - Context or initial state for the base component.
     */
    function CheckboxBase(props, context) {
        var _this = _super.call(this, props, context) || this;
        _this._checkBox = react__WEBPACK_IMPORTED_MODULE_1__["createRef"]();
        _this._renderContent = function (checked, indeterminate, keytipAttributes) {
            if (keytipAttributes === void 0) { keytipAttributes = {}; }
            var _a = _this.props, disabled = _a.disabled, inputProps = _a.inputProps, name = _a.name, ariaLabel = _a.ariaLabel, ariaLabelledBy = _a.ariaLabelledBy, ariaDescribedBy = _a.ariaDescribedBy, _b = _a.onRenderLabel, onRenderLabel = _b === void 0 ? _this._onRenderLabel : _b, checkmarkIconProps = _a.checkmarkIconProps, ariaPositionInSet = _a.ariaPositionInSet, ariaSetSize = _a.ariaSetSize, title = _a.title, label = _a.label;
            return (react__WEBPACK_IMPORTED_MODULE_1__["createElement"]("div", { className: _this._classNames.root, title: title },
                react__WEBPACK_IMPORTED_MODULE_1__["createElement"](_Utilities__WEBPACK_IMPORTED_MODULE_3__["FocusRects"], null),
                react__WEBPACK_IMPORTED_MODULE_1__["createElement"]("input", Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__assign"])({ type: "checkbox" }, inputProps, { "data-ktp-execute-target": keytipAttributes['data-ktp-execute-target'], checked: checked, disabled: disabled, className: _this._classNames.input, ref: _this._checkBox, name: name, id: _this._id, title: title, onChange: _this._onChange, onFocus: _this._onFocus, onBlur: _this._onBlur, "aria-disabled": disabled, "aria-label": ariaLabel || label, "aria-labelledby": ariaLabelledBy, "aria-describedby": Object(_Utilities__WEBPACK_IMPORTED_MODULE_4__["mergeAriaAttributeValues"])(ariaDescribedBy, keytipAttributes['aria-describedby']), "aria-posinset": ariaPositionInSet, "aria-setsize": ariaSetSize, "aria-checked": indeterminate ? 'mixed' : checked ? 'true' : 'false' })),
                react__WEBPACK_IMPORTED_MODULE_1__["createElement"]("label", { className: _this._classNames.label, htmlFor: _this._id },
                    react__WEBPACK_IMPORTED_MODULE_1__["createElement"]("div", { className: _this._classNames.checkbox, "data-ktp-target": keytipAttributes['data-ktp-target'] },
                        react__WEBPACK_IMPORTED_MODULE_1__["createElement"](_Icon__WEBPACK_IMPORTED_MODULE_8__["Icon"], Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__assign"])({ iconName: "CheckMark" }, checkmarkIconProps, { className: _this._classNames.checkmark }))),
                    onRenderLabel(_this.props, _this._onRenderLabel))));
        };
        _this._onFocus = function (ev) {
            var inputProps = _this.props.inputProps;
            if (inputProps && inputProps.onFocus) {
                inputProps.onFocus(ev);
            }
        };
        _this._onBlur = function (ev) {
            var inputProps = _this.props.inputProps;
            if (inputProps && inputProps.onBlur) {
                inputProps.onBlur(ev);
            }
        };
        _this._onChange = function (ev) {
            var onChange = _this.props.onChange;
            var _a = _this.state, isChecked = _a.isChecked, isIndeterminate = _a.isIndeterminate;
            if (!isIndeterminate) {
                if (onChange) {
                    onChange(ev, !isChecked);
                }
                if (_this.props.checked === undefined) {
                    _this.setState({ isChecked: !isChecked });
                }
            }
            else {
                // If indeterminate, clicking the checkbox *only* removes the indeterminate state (or if
                // controlled, lets the consumer know to change it by calling onChange). It doesn't
                // change the checked state.
                if (onChange) {
                    onChange(ev, isChecked);
                }
                if (_this.props.indeterminate === undefined) {
                    _this.setState({ isIndeterminate: false });
                }
            }
        };
        _this._onRenderLabel = function (props) {
            var label = props.label, title = props.title;
            return label ? (react__WEBPACK_IMPORTED_MODULE_1__["createElement"]("span", { "aria-hidden": "true", className: _this._classNames.text, title: title }, label)) : null;
        };
        Object(_Utilities__WEBPACK_IMPORTED_MODULE_5__["initializeComponentRef"])(_this);
        if (true) {
            Object(_Utilities__WEBPACK_IMPORTED_MODULE_6__["warnMutuallyExclusive"])('Checkbox', props, {
                checked: 'defaultChecked',
                indeterminate: 'defaultIndeterminate',
            });
        }
        _this._id = _this.props.id || Object(_Utilities__WEBPACK_IMPORTED_MODULE_7__["getId"])('checkbox-');
        _this.state = {
            isChecked: !!(props.checked !== undefined ? props.checked : props.defaultChecked),
            isIndeterminate: !!(props.indeterminate !== undefined ? props.indeterminate : props.defaultIndeterminate),
        };
        return _this;
    }
    CheckboxBase.getDerivedStateFromProps = function (nextProps, prevState) {
        var stateUpdate = {};
        if (nextProps.indeterminate !== undefined) {
            stateUpdate.isIndeterminate = !!nextProps.indeterminate;
        }
        if (nextProps.checked !== undefined) {
            stateUpdate.isChecked = !!nextProps.checked;
        }
        return Object.keys(stateUpdate).length ? stateUpdate : null;
    };
    /**
     * Render the Checkbox based on passed props
     */
    CheckboxBase.prototype.render = function () {
        var _this = this;
        var _a = this.props, className = _a.className, disabled = _a.disabled, boxSide = _a.boxSide, theme = _a.theme, styles = _a.styles, _b = _a.onRenderLabel, onRenderLabel = _b === void 0 ? this._onRenderLabel : _b, keytipProps = _a.keytipProps;
        var _c = this.state, isChecked = _c.isChecked, isIndeterminate = _c.isIndeterminate;
        this._classNames = getClassNames(styles, {
            theme: theme,
            className: className,
            disabled: disabled,
            indeterminate: isIndeterminate,
            checked: isChecked,
            reversed: boxSide !== 'start',
            isUsingCustomLabelRender: onRenderLabel !== this._onRenderLabel,
        });
        if (keytipProps) {
            return (react__WEBPACK_IMPORTED_MODULE_1__["createElement"](_KeytipData__WEBPACK_IMPORTED_MODULE_9__["KeytipData"], { keytipProps: keytipProps, disabled: disabled }, function (keytipAttributes) { return _this._renderContent(isChecked, isIndeterminate, keytipAttributes); }));
        }
        return this._renderContent(isChecked, isIndeterminate);
    };
    Object.defineProperty(CheckboxBase.prototype, "indeterminate", {
        get: function () {
            return !!this.state.isIndeterminate;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(CheckboxBase.prototype, "checked", {
        get: function () {
            return !!this.state.isChecked;
        },
        enumerable: true,
        configurable: true
    });
    CheckboxBase.prototype.focus = function () {
        if (this._checkBox.current) {
            this._checkBox.current.focus();
        }
    };
    CheckboxBase.defaultProps = {
        boxSide: 'start',
    };
    return CheckboxBase;
}(react__WEBPACK_IMPORTED_MODULE_1__["Component"]));



/***/ }),

/***/ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/office-ui-fabric-react/7.121.6_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/office-ui-fabric-react/lib/components/Checkbox/Checkbox.js":
/*!******************************************************************************************************************************************************************************************************************************************************!*\
  !*** C:/Users/riarya/Desktop/www/office-bohemia/common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/office-ui-fabric-react/7.121.6_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/office-ui-fabric-react/lib/components/Checkbox/Checkbox.js ***!
  \******************************************************************************************************************************************************************************************************************************************************/
/*! exports provided: Checkbox */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Checkbox", function() { return Checkbox; });
/* harmony import */ var _Utilities__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../../Utilities */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/styled.js");
/* harmony import */ var _Checkbox_base__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./Checkbox.base */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/office-ui-fabric-react/7.121.6_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/office-ui-fabric-react/lib/components/Checkbox/Checkbox.base.js");
/* harmony import */ var _Checkbox_styles__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./Checkbox.styles */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/office-ui-fabric-react/7.121.6_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/office-ui-fabric-react/lib/components/Checkbox/Checkbox.styles.js");



var Checkbox = Object(_Utilities__WEBPACK_IMPORTED_MODULE_0__["styled"])(_Checkbox_base__WEBPACK_IMPORTED_MODULE_1__["CheckboxBase"], _Checkbox_styles__WEBPACK_IMPORTED_MODULE_2__["getStyles"], undefined, { scope: 'Checkbox' });


/***/ }),

/***/ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/office-ui-fabric-react/7.121.6_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/office-ui-fabric-react/lib/components/Checkbox/Checkbox.styles.js":
/*!*************************************************************************************************************************************************************************************************************************************************************!*\
  !*** C:/Users/riarya/Desktop/www/office-bohemia/common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/office-ui-fabric-react/7.121.6_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/office-ui-fabric-react/lib/components/Checkbox/Checkbox.styles.js ***!
  \*************************************************************************************************************************************************************************************************************************************************************/
/*! exports provided: getStyles */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getStyles", function() { return getStyles; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/tslib/1.13.0/node_modules/tslib/tslib.es6.js");
/* harmony import */ var _Styling__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../../Styling */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/styling/7.13.5_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/styling/lib/index.js");
/* harmony import */ var _Utilities__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../../Utilities */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/setFocusVisibility.js");



var GlobalClassNames = {
    root: 'ms-Checkbox',
    label: 'ms-Checkbox-label',
    checkbox: 'ms-Checkbox-checkbox',
    checkmark: 'ms-Checkbox-checkmark',
    text: 'ms-Checkbox-text',
};
var MS_CHECKBOX_LABEL_SIZE = '20px';
var MS_CHECKBOX_TRANSITION_DURATION = '200ms';
var MS_CHECKBOX_TRANSITION_TIMING = 'cubic-bezier(.4, 0, .23, 1)';
var getStyles = function (props) {
    var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m, _o, _p, _q, _r, _s, _t;
    var className = props.className, theme = props.theme, reversed = props.reversed, checked = props.checked, disabled = props.disabled, isUsingCustomLabelRender = props.isUsingCustomLabelRender, indeterminate = props.indeterminate;
    var semanticColors = theme.semanticColors, effects = theme.effects, palette = theme.palette, fonts = theme.fonts;
    var classNames = Object(_Styling__WEBPACK_IMPORTED_MODULE_1__["getGlobalClassNames"])(GlobalClassNames, theme);
    var checkmarkFontColor = semanticColors.inputForegroundChecked;
    // TODO: after updating the semanticColors slots mapping this needs to be semanticColors.inputBorder
    var checkmarkFontColorHovered = palette.neutralSecondary;
    // TODO: after updating the semanticColors slots mapping this needs to be semanticColors.smallInputBorder
    var checkboxBorderColor = palette.neutralPrimary;
    var checkboxBorderIndeterminateColor = semanticColors.inputBackgroundChecked;
    var checkboxBorderColorChecked = semanticColors.inputBackgroundChecked;
    var checkboxBorderColorDisabled = semanticColors.disabledBodySubtext;
    var checkboxBorderHoveredColor = semanticColors.inputBorderHovered;
    var checkboxBorderIndeterminateHoveredColor = semanticColors.inputBackgroundCheckedHovered;
    var checkboxBackgroundChecked = semanticColors.inputBackgroundChecked;
    // TODO: after updating the semanticColors slots mapping the following 2 tokens need to be
    // semanticColors.inputBackgroundCheckedHovered
    var checkboxBackgroundCheckedHovered = semanticColors.inputBackgroundCheckedHovered;
    var checkboxBorderColorCheckedHovered = semanticColors.inputBackgroundCheckedHovered;
    var checkboxHoveredTextColor = semanticColors.inputTextHovered;
    var checkboxBackgroundDisabledChecked = semanticColors.disabledBodySubtext;
    var checkboxTextColor = semanticColors.bodyText;
    var checkboxTextColorDisabled = semanticColors.disabledText;
    var indeterminateDotStyles = [
        {
            content: '""',
            borderRadius: effects.roundedCorner2,
            position: 'absolute',
            width: 10,
            height: 10,
            top: 4,
            left: 4,
            boxSizing: 'border-box',
            borderWidth: 5,
            borderStyle: 'solid',
            borderColor: disabled ? checkboxBorderColorDisabled : checkboxBorderIndeterminateColor,
            transitionProperty: 'border-width, border, border-color',
            transitionDuration: MS_CHECKBOX_TRANSITION_DURATION,
            transitionTimingFunction: MS_CHECKBOX_TRANSITION_TIMING,
            selectors: (_a = {},
                _a[_Styling__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]] = {
                    borderColor: 'WindowText',
                },
                _a),
        },
    ];
    return {
        root: [
            classNames.root,
            {
                position: 'relative',
                display: 'flex',
            },
            reversed && 'reversed',
            checked && 'is-checked',
            !disabled && 'is-enabled',
            disabled && 'is-disabled',
            !disabled && [
                !checked && {
                    selectors: (_b = {},
                        _b[":hover ." + classNames.checkbox] = {
                            borderColor: checkboxBorderHoveredColor,
                            selectors: (_c = {},
                                _c[_Styling__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]] = {
                                    borderColor: 'Highlight',
                                },
                                _c),
                        },
                        _b[":focus ." + classNames.checkbox] = { borderColor: checkboxBorderHoveredColor },
                        _b[":hover ." + classNames.checkmark] = {
                            color: checkmarkFontColorHovered,
                            opacity: '1',
                            selectors: (_d = {},
                                _d[_Styling__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]] = {
                                    color: 'Highlight',
                                },
                                _d),
                        },
                        _b),
                },
                checked &&
                    !indeterminate && {
                    selectors: (_e = {},
                        _e[":hover ." + classNames.checkbox] = {
                            background: checkboxBackgroundCheckedHovered,
                            borderColor: checkboxBorderColorCheckedHovered,
                        },
                        _e[":focus ." + classNames.checkbox] = {
                            background: checkboxBackgroundCheckedHovered,
                            borderColor: checkboxBorderColorCheckedHovered,
                        },
                        _e[_Styling__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]] = {
                            selectors: (_f = {},
                                _f[":hover ." + classNames.checkbox] = {
                                    background: 'Window',
                                    borderColor: 'Highlight',
                                },
                                _f[":focus ." + classNames.checkbox] = {
                                    background: 'Highlight',
                                },
                                _f[":focus:hover ." + classNames.checkbox] = {
                                    background: 'Highlight',
                                },
                                _f[":focus:hover ." + classNames.checkmark] = {
                                    color: 'Window',
                                },
                                _f[":hover ." + classNames.checkmark] = {
                                    color: 'Highlight',
                                },
                                _f),
                        },
                        _e),
                },
                indeterminate && {
                    selectors: (_g = {},
                        _g[":hover ." + classNames.checkbox + ", :hover ." + classNames.checkbox + ":after"] = {
                            borderColor: checkboxBorderIndeterminateHoveredColor,
                            selectors: (_h = {},
                                _h[_Styling__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]] = {
                                    borderColor: 'WindowText',
                                },
                                _h),
                        },
                        _g[":focus ." + classNames.checkbox] = {
                            borderColor: checkboxBorderIndeterminateHoveredColor,
                        },
                        _g[":hover ." + classNames.checkmark] = {
                            opacity: '0',
                        },
                        _g),
                },
                {
                    selectors: (_j = {},
                        _j[":hover ." + classNames.text + ", :focus ." + classNames.text] = {
                            color: checkboxHoveredTextColor,
                            selectors: (_k = {},
                                _k[_Styling__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]] = {
                                    color: disabled ? 'GrayText' : 'WindowText',
                                },
                                _k),
                        },
                        _j),
                },
            ],
            className,
        ],
        input: {
            position: 'absolute',
            background: 'none',
            opacity: 0,
            selectors: (_l = {},
                _l["." + _Utilities__WEBPACK_IMPORTED_MODULE_2__["IsFocusVisibleClassName"] + " &:focus + label::before"] = {
                    outline: '1px solid ' + theme.palette.neutralSecondary,
                    outlineOffset: '2px',
                    selectors: (_m = {},
                        _m[_Styling__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]] = {
                            outline: '1px solid ActiveBorder',
                        },
                        _m),
                },
                _l),
        },
        label: [
            classNames.label,
            theme.fonts.medium,
            {
                display: 'flex',
                alignItems: isUsingCustomLabelRender ? 'center' : 'flex-start',
                cursor: disabled ? 'default' : 'pointer',
                position: 'relative',
                userSelect: 'none',
            },
            reversed && {
                flexDirection: 'row-reverse',
                justifyContent: 'flex-end',
            },
            {
                selectors: {
                    '&::before': {
                        position: 'absolute',
                        left: 0,
                        right: 0,
                        top: 0,
                        bottom: 0,
                        content: '""',
                        pointerEvents: 'none',
                    },
                },
            },
        ],
        checkbox: [
            classNames.checkbox,
            {
                position: 'relative',
                display: 'flex',
                flexShrink: 0,
                alignItems: 'center',
                justifyContent: 'center',
                height: MS_CHECKBOX_LABEL_SIZE,
                width: MS_CHECKBOX_LABEL_SIZE,
                border: "1px solid " + checkboxBorderColor,
                borderRadius: effects.roundedCorner2,
                boxSizing: 'border-box',
                transitionProperty: 'background, border, border-color',
                transitionDuration: MS_CHECKBOX_TRANSITION_DURATION,
                transitionTimingFunction: MS_CHECKBOX_TRANSITION_TIMING,
                /* in case the icon is bigger than the box */
                overflow: 'hidden',
                selectors: Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__assign"])((_o = { ':after': indeterminate ? indeterminateDotStyles : null }, _o[_Styling__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]] = {
                    borderColor: 'WindowText',
                }, _o), Object(_Styling__WEBPACK_IMPORTED_MODULE_1__["getEdgeChromiumNoHighContrastAdjustSelector"])()),
            },
            indeterminate && {
                borderColor: checkboxBorderIndeterminateColor,
            },
            !reversed
                ? // This margin on the checkbox is for backwards compat. Notably it has the effect where a customRender
                    // is used, there will be only a 4px margin from checkbox to label. The label by default would have
                    // another 4px margin for a total of 8px margin between checkbox and label. We don't combine the two
                    // (and move it into the text) to not incur a breaking change for everyone using custom render atm.
                    {
                        marginRight: 4,
                    }
                : {
                    marginLeft: 4,
                },
            !disabled &&
                !indeterminate &&
                checked && {
                background: checkboxBackgroundChecked,
                borderColor: checkboxBorderColorChecked,
                selectors: (_p = {},
                    _p[_Styling__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]] = {
                        background: 'Highlight',
                        borderColor: 'Highlight',
                    },
                    _p),
            },
            disabled && {
                borderColor: checkboxBorderColorDisabled,
                selectors: (_q = {},
                    _q[_Styling__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]] = {
                        borderColor: 'GrayText',
                    },
                    _q),
            },
            checked &&
                disabled && {
                background: checkboxBackgroundDisabledChecked,
                borderColor: checkboxBorderColorDisabled,
                selectors: (_r = {},
                    _r[_Styling__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]] = {
                        background: 'Window',
                    },
                    _r),
            },
        ],
        checkmark: [
            classNames.checkmark,
            {
                opacity: checked ? '1' : '0',
                color: checkmarkFontColor,
                selectors: (_s = {},
                    _s[_Styling__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]] = {
                        color: disabled ? 'GrayText' : 'Window',
                        MsHighContrastAdjust: 'none',
                    },
                    _s),
            },
        ],
        text: [
            classNames.text,
            {
                color: disabled ? checkboxTextColorDisabled : checkboxTextColor,
                fontSize: fonts.medium.fontSize,
                lineHeight: '20px',
                selectors: Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__assign"])((_t = {}, _t[_Styling__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]] = {
                    color: disabled ? 'GrayText' : 'WindowText',
                }, _t), Object(_Styling__WEBPACK_IMPORTED_MODULE_1__["getEdgeChromiumNoHighContrastAdjustSelector"])()),
            },
            !reversed
                ? {
                    marginLeft: 4,
                }
                : {
                    marginRight: 4,
                },
        ],
    };
};


/***/ }),

/***/ "../office-bohemia-ux/lib/utilities/nodeContainsTarget.js":
/*!****************************************************************!*\
  !*** ../office-bohemia-ux/lib/utilities/nodeContainsTarget.js ***!
  \****************************************************************/
/*! exports provided: nodeContainsTarget */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "nodeContainsTarget", function() { return nodeContainsTarget; });
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/dom/setPortalAttribute.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/dom/getVirtualParent.js");

/**
 * Checks if the `target` is contained within the `node` DOM tree structure.
 * The utility is considering also elements from floating UI as being part of the tree
 * **only** if they are linked to a node which is part of that DOM hierarchy already.
 *
 * For our current floating UI implementation we are using *Callout* controls from *office-ui-fabric-react* library.
 * This control is built using React Portals (https://reactjs.org/docs/portals.html). A portal has two linked elements,
 * one within the same DOM hierarchy and the other one outside. The one outside is marked with a special attribute
 * and is called a `virtual child`. The other one is called `virtual parent`.
 *
 * In order for a `target` element under the *virtual child* DOM hierarchy to be considered a descendant of the `node`,
 * the `node` needs to contain the *virtual parent*.
 * @param node The root of the DOM hierarchy under we search for the `target`
 * @param target The element we are looking for.
 */
function nodeContainsTarget(node, target) {
    // First we check if the target is an ancestor of the given node.
    if (node.contains(target)) {
        return true;
    }
    // Safety check to ensure the cast to HTMLElement below.
    if (target.nodeType !== document.ELEMENT_NODE) {
        return false;
    }
    let current = target;
    // Try to find the portal element.
    while (current && !current.hasAttribute(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["DATA_PORTAL_ATTRIBUTE"])) {
        current = current.parentElement;
    }
    if (!current) {
        return false;
    }
    // Try to get the element from the other side of the portal (called virtual parent).
    const virtualParent = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["getVirtualParent"])(current);
    if (!virtualParent) {
        return false;
    }
    return node.contains(virtualParent);
}
// cSpell:ignore Callout


/***/ }),

/***/ "../tablero/lib/commanding/getCommands.js":
/*!************************************************!*\
  !*** ../tablero/lib/commanding/getCommands.js ***!
  \************************************************/
/*! exports provided: getCommandingSurfaceCommands */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getCommandingSurfaceCommands", function() { return getCommandingSurfaceCommands; });
const getCommandingSurfaceCommands = (context, commandFactories) => {
    let commandGroups = [];
    commandFactories.forEach((factoryGroup) => {
        const commandGroup = [];
        factoryGroup.forEach((factory) => {
            const command = factory(context);
            if (command !== undefined) {
                commandGroup.push(command);
            }
        });
        // Don't create a new group if this group has no commands.
        if (commandGroup.length > 0) {
            commandGroups.push(commandGroup);
        }
    });
    return commandGroups;
};


/***/ }),

/***/ "../tablero/lib/configuration/Configuration.js":
/*!*****************************************************!*\
  !*** ../tablero/lib/configuration/Configuration.js ***!
  \*****************************************************/
/*! exports provided: Configuration */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Configuration", function() { return Configuration; });
/**
 * Enumeration of known tablero configurations.
 * Except for `custom`, there is a matching stored template for each other configuration
 */
var Configuration;
(function (Configuration) {
    Configuration["default"] = "default";
    Configuration["actionItemManager"] = "actionItemManager";
    Configuration["richTextTablero"] = "richTextTablero";
    Configuration["tasks"] = "tasks";
})(Configuration || (Configuration = {}));


/***/ }),

/***/ "../tablero/lib/configuration/ConfigurationUtils.js":
/*!**********************************************************!*\
  !*** ../tablero/lib/configuration/ConfigurationUtils.js ***!
  \**********************************************************/
/*! exports provided: ConfigurationFactory, getConfigurationFriendlyName, getConfigurationPreset, createAndGetTableroDocumentId */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ConfigurationFactory", function() { return ConfigurationFactory; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getConfigurationFriendlyName", function() { return getConfigurationFriendlyName; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getConfigurationPreset", function() { return getConfigurationPreset; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "createAndGetTableroDocumentId", function() { return createAndGetTableroDocumentId; });
/* harmony import */ var _Configuration__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./Configuration */ "../tablero/lib/configuration/Configuration.js");
/* harmony import */ var _actionItemManager__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./actionItemManager */ "../tablero/lib/configuration/actionItemManager.js");
/* harmony import */ var _defaultTablero__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./defaultTablero */ "../tablero/lib/configuration/defaultTablero.js");
/* harmony import */ var _richTextTablero__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./richTextTablero */ "../tablero/lib/configuration/richTextTablero.js");
/* harmony import */ var _tasks__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./tasks */ "../tablero/lib/configuration/tasks.js");
/* harmony import */ var _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! @fluidx/office-fluid-types */ "../office-fluid-types/lib/ComponentTypes.js");
/* harmony import */ var uuid_v4__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! uuid/v4 */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/uuid/3.4.0/node_modules/uuid/v4.js");
/* harmony import */ var uuid_v4__WEBPACK_IMPORTED_MODULE_6___default = /*#__PURE__*/__webpack_require__.n(uuid_v4__WEBPACK_IMPORTED_MODULE_6__);







class ConfigurationFactory {
    constructor() {
        this.create = (configuration) => {
            let { columnTypes, numOfRows, componentFriendlyName, configurationType } = configuration;
            // Setting up the default rows and columns if the number of rows and column types are provided
            if (numOfRows && columnTypes && columnTypes.length > 0) {
                const columns = Object(_defaultTablero__WEBPACK_IMPORTED_MODULE_2__["getDefaultTableroColumns"])(columnTypes);
                return {
                    configurationType,
                    createType: 'newTablero',
                    columns,
                    rows: Object(_defaultTablero__WEBPACK_IMPORTED_MODULE_2__["getDefaultTableroRows"])(numOfRows, columns),
                    componentFriendlyName: componentFriendlyName || getConfigurationFriendlyName(configurationType)
                };
            }
            switch (configuration.configurationType) {
                case _Configuration__WEBPACK_IMPORTED_MODULE_0__["Configuration"].actionItemManager:
                    return _actionItemManager__WEBPACK_IMPORTED_MODULE_1__["actionItemManager"];
                case _Configuration__WEBPACK_IMPORTED_MODULE_0__["Configuration"].richTextTablero:
                    return _richTextTablero__WEBPACK_IMPORTED_MODULE_3__["richTextTablero"];
                case _Configuration__WEBPACK_IMPORTED_MODULE_0__["Configuration"].tasks:
                    return _tasks__WEBPACK_IMPORTED_MODULE_4__["tasks"];
                default:
                    return _defaultTablero__WEBPACK_IMPORTED_MODULE_2__["defaultTablero"];
            }
        };
    }
}
const getConfigurationFriendlyName = (configurationType) => {
    switch (configurationType) {
        case _Configuration__WEBPACK_IMPORTED_MODULE_0__["Configuration"].actionItemManager:
            return _actionItemManager__WEBPACK_IMPORTED_MODULE_1__["actionItemManager"].componentFriendlyName;
        case _Configuration__WEBPACK_IMPORTED_MODULE_0__["Configuration"].richTextTablero:
            return _richTextTablero__WEBPACK_IMPORTED_MODULE_3__["richTextTablero"].componentFriendlyName;
        case _Configuration__WEBPACK_IMPORTED_MODULE_0__["Configuration"].tasks:
            return _tasks__WEBPACK_IMPORTED_MODULE_4__["tasks"].componentFriendlyName;
        default:
            return _defaultTablero__WEBPACK_IMPORTED_MODULE_2__["defaultTablero"].componentFriendlyName;
    }
};
const getConfigurationPreset = (configurationType, presetName) => {
    switch (configurationType) {
        case _Configuration__WEBPACK_IMPORTED_MODULE_0__["Configuration"].actionItemManager:
            return _actionItemManager__WEBPACK_IMPORTED_MODULE_1__["actionItemManagerPresets"].get(presetName);
        case _Configuration__WEBPACK_IMPORTED_MODULE_0__["Configuration"].richTextTablero:
            return _richTextTablero__WEBPACK_IMPORTED_MODULE_3__["richTextTableroPresets"].get(presetName);
        case _Configuration__WEBPACK_IMPORTED_MODULE_0__["Configuration"].tasks:
            return _tasks__WEBPACK_IMPORTED_MODULE_4__["tasksPresets"].get(presetName);
        default:
            return _defaultTablero__WEBPACK_IMPORTED_MODULE_2__["defaultTableroPresets"].get(presetName);
    }
};
const createAndGetTableroDocumentId = async (config, context) => {
    const tableroDocumentId = uuid_v4__WEBPACK_IMPORTED_MODULE_6__();
    const runtime = await context.createComponent(tableroDocumentId, _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_5__["TableroDocumentType"], config);
    runtime.bindToContext();
    return Promise.resolve(tableroDocumentId);
};


/***/ }),

/***/ "../tablero/lib/configuration/Presets.js":
/*!***********************************************!*\
  !*** ../tablero/lib/configuration/Presets.js ***!
  \***********************************************/
/*! exports provided: PresetName, defaultPresetsMap, overrideDefaultPresets */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "PresetName", function() { return PresetName; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "defaultPresetsMap", function() { return defaultPresetsMap; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "overrideDefaultPresets", function() { return overrideDefaultPresets; });
var PresetName;
(function (PresetName) {
    PresetName[PresetName["DisableInsertColumns"] = 0] = "DisableInsertColumns";
    PresetName[PresetName["DisableInsertRows"] = 1] = "DisableInsertRows";
    PresetName[PresetName["DisableDeleteColumns"] = 2] = "DisableDeleteColumns";
    PresetName[PresetName["DisableDeleteRows"] = 3] = "DisableDeleteRows";
    PresetName[PresetName["DisableRenameColumns"] = 4] = "DisableRenameColumns";
    PresetName[PresetName["DisableSort"] = 5] = "DisableSort";
    PresetName[PresetName["DisableResizeColumnFromCmdMenu"] = 6] = "DisableResizeColumnFromCmdMenu";
    PresetName[PresetName["EnableAddRowButton"] = 7] = "EnableAddRowButton";
    /** Deprecated: TODO: Task 3749578: Remove presets persistence back compatibility */
    PresetName[PresetName["EnableInsertColumns"] = 8] = "EnableInsertColumns";
})(PresetName || (PresetName = {}));
/**
 * Presets determining functionality / feature availability in a tablero configuration
 * Add a configuration property name in `PresetName` and a
 * default option in this settings map
 */
const defaultPresetsMap = new Map([
    [PresetName.DisableInsertColumns, false],
    [PresetName.DisableInsertRows, false],
    [PresetName.DisableDeleteColumns, false],
    [PresetName.DisableDeleteRows, false],
    [PresetName.DisableRenameColumns, false],
    [PresetName.DisableSort, false],
    [PresetName.DisableResizeColumnFromCmdMenu, false],
    [PresetName.EnableAddRowButton, false]
]);
const overrideDefaultPresets = (overrides) => {
    if (overrides !== undefined) {
        defaultPresetsMap.forEach((preset, presetName) => {
            if (!overrides.has(presetName)) {
                overrides.set(presetName, preset);
            }
        });
        return overrides;
    }
    return defaultPresetsMap;
};


/***/ }),

/***/ "../tablero/lib/configuration/SettingsValueTypes.js":
/*!**********************************************************!*\
  !*** ../tablero/lib/configuration/SettingsValueTypes.js ***!
  \**********************************************************/
/*! exports provided: SettingsValueTypes */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SettingsValueTypes", function() { return SettingsValueTypes; });
const packageName = 'tablero';
/**
 * Mapping names of settings values used across this package.
 */
const SettingsValueTypes = {
    /**
     * Flag name for determining if should nest any component's layout.
     */
    nestAnyLayout: `${packageName}.nestAnyLayout`,
    /**
     * Flag name for determining if tablero should have side margins
     */
    sideMargin: `${packageName}.sideMargin`,
    /**
     * Flag name for determining if tablero row and column grabbers should be enabled
     */
    grabbers: `${packageName}.grabbers`,
    /**
     * Flag name for determining if drag and drop should be enabled
     */
    dragAndDrop: `${packageName}.dragAndDrop`,
    /**
     * Flag name for determining if resize column menu item in commanding menu should be enabled.
     */
    resizeColumnFromCmdMenu: `${packageName}.resizeColumnFromCmdMenu`,
    /**
     * Flag name for determining if Rich Text Table should be enabled
     */
    richTextTable: `${packageName}.richTextTable`
};


/***/ }),

/***/ "../tablero/lib/configuration/actionItemManager.js":
/*!*********************************************************!*\
  !*** ../tablero/lib/configuration/actionItemManager.js ***!
  \*********************************************************/
/*! exports provided: actionItemManager, actionItemManagerPresets */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "actionItemManager", function() { return actionItemManager; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "actionItemManagerPresets", function() { return actionItemManagerPresets; });
/* harmony import */ var _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../view/table/Table.props */ "../tablero/lib/view/table/Table.props.js");
/* harmony import */ var _utils__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./utils */ "../tablero/lib/configuration/utils.js");
/* harmony import */ var _Presets__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./Presets */ "../tablero/lib/configuration/Presets.js");
/* harmony import */ var _Configuration__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./Configuration */ "../tablero/lib/configuration/Configuration.js");




const actionItemManagerColumns = [
    { title: 'Done', id: 'initialColumn-0', width: 106, type: _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].Checkbox },
    { title: 'Task Name', id: 'initialColumn-1', width: 179, type: _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].Text },
    { title: 'Assignee', id: 'initialColumn-2', width: 174, type: _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].Text },
    { title: 'Due Date', id: 'initialColumn-3', width: 205, type: _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].Date },
    { title: 'Notes', id: 'initialColumn-4', width: 167, type: _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].Text }
];
const actionItemManagerRows = [
    Object(_utils__WEBPACK_IMPORTED_MODULE_1__["getTableRow"])('initialRow-0', actionItemManagerColumns),
    Object(_utils__WEBPACK_IMPORTED_MODULE_1__["getTableRow"])('initialRow-1', actionItemManagerColumns),
    Object(_utils__WEBPACK_IMPORTED_MODULE_1__["getTableRow"])('initialRow-2', actionItemManagerColumns)
];
const presetOverrides = new Map([
    [_Presets__WEBPACK_IMPORTED_MODULE_2__["PresetName"].DisableInsertColumns, true],
    [_Presets__WEBPACK_IMPORTED_MODULE_2__["PresetName"].DisableDeleteColumns, true],
    [_Presets__WEBPACK_IMPORTED_MODULE_2__["PresetName"].DisableRenameColumns, true]
]);
const actionItemManager = {
    configurationType: _Configuration__WEBPACK_IMPORTED_MODULE_3__["Configuration"].actionItemManager,
    createType: 'newTablero',
    componentFriendlyName: 'Action Items',
    columns: actionItemManagerColumns,
    rows: actionItemManagerRows
};
const actionItemManagerPresets = Object(_Presets__WEBPACK_IMPORTED_MODULE_2__["overrideDefaultPresets"])(presetOverrides);


/***/ }),

/***/ "../tablero/lib/configuration/defaultTablero.js":
/*!******************************************************!*\
  !*** ../tablero/lib/configuration/defaultTablero.js ***!
  \******************************************************/
/*! exports provided: getDefaultTableroColumns, getDefaultTableroRows, defaultTablero, defaultTableroPresets */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getDefaultTableroColumns", function() { return getDefaultTableroColumns; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getDefaultTableroRows", function() { return getDefaultTableroRows; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "defaultTablero", function() { return defaultTablero; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "defaultTableroPresets", function() { return defaultTableroPresets; });
/* harmony import */ var _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../view/table/Table.props */ "../tablero/lib/view/table/Table.props.js");
/* harmony import */ var _utils__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./utils */ "../tablero/lib/configuration/utils.js");
/* harmony import */ var _Presets__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./Presets */ "../tablero/lib/configuration/Presets.js");
/* harmony import */ var _Configuration__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./Configuration */ "../tablero/lib/configuration/Configuration.js");
/* harmony import */ var _view_StylingConstants__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../view/StylingConstants */ "../tablero/lib/view/StylingConstants.js");





const defaultColumnTypes = [_view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].Text, _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].Text, _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].Text, _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].Text];
const defaultNumOfRows = 3;
const getDefaultTableroColumns = (columnTypes) => {
    return Object(_utils__WEBPACK_IMPORTED_MODULE_1__["getTableroColumns"])(columnTypes, _view_StylingConstants__WEBPACK_IMPORTED_MODULE_4__["defaultColumnWidth"]);
};
const getDefaultTableroRows = (nRows, columns) => {
    return Object(_utils__WEBPACK_IMPORTED_MODULE_1__["getTableroRows"])(nRows, columns);
};
const defaultTableroColumns = getDefaultTableroColumns(defaultColumnTypes);
const defaultTableroRows = getDefaultTableroRows(defaultNumOfRows, defaultTableroColumns);
const defaultTablero = {
    configurationType: _Configuration__WEBPACK_IMPORTED_MODULE_3__["Configuration"].default,
    createType: 'newTablero',
    componentFriendlyName: 'Fluid Table',
    columns: defaultTableroColumns,
    rows: defaultTableroRows
};
const defaultTableroPresets = _Presets__WEBPACK_IMPORTED_MODULE_2__["defaultPresetsMap"];


/***/ }),

/***/ "../tablero/lib/configuration/richTextTablero.js":
/*!*******************************************************!*\
  !*** ../tablero/lib/configuration/richTextTablero.js ***!
  \*******************************************************/
/*! exports provided: richTextTablero, richTextTableroPresets */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "richTextTablero", function() { return richTextTablero; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "richTextTableroPresets", function() { return richTextTableroPresets; });
/* harmony import */ var _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../view/table/Table.props */ "../tablero/lib/view/table/Table.props.js");
/* harmony import */ var _utils__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./utils */ "../tablero/lib/configuration/utils.js");
/* harmony import */ var _Configuration__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./Configuration */ "../tablero/lib/configuration/Configuration.js");
/* harmony import */ var _Presets__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./Presets */ "../tablero/lib/configuration/Presets.js");
/* harmony import */ var _view_StylingConstants__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../view/StylingConstants */ "../tablero/lib/view/StylingConstants.js");





const columnTypes = [_view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].RichText, _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].RichText, _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].RichText, _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].RichText];
const numOfRows = 3;
const richTextColumns = Object(_utils__WEBPACK_IMPORTED_MODULE_1__["getTableroColumns"])(columnTypes, _view_StylingConstants__WEBPACK_IMPORTED_MODULE_4__["defaultColumnWidth"]);
const richTextTableroRows = Object(_utils__WEBPACK_IMPORTED_MODULE_1__["getTableroRows"])(numOfRows, richTextColumns);
const richTextTablero = {
    configurationType: _Configuration__WEBPACK_IMPORTED_MODULE_2__["Configuration"].richTextTablero,
    createType: 'newTablero',
    componentFriendlyName: 'Fluid Table',
    columns: richTextColumns,
    rows: richTextTableroRows
};
const richTextTableroPresets = _Presets__WEBPACK_IMPORTED_MODULE_3__["defaultPresetsMap"];


/***/ }),

/***/ "../tablero/lib/configuration/tasks.js":
/*!*********************************************!*\
  !*** ../tablero/lib/configuration/tasks.js ***!
  \*********************************************/
/*! exports provided: defaultNumbersRows, tasks, tasksPresets */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "defaultNumbersRows", function() { return defaultNumbersRows; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tasks", function() { return tasks; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tasksPresets", function() { return tasksPresets; });
/* harmony import */ var _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../view/table/Table.props */ "../tablero/lib/view/table/Table.props.js");
/* harmony import */ var _utils__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./utils */ "../tablero/lib/configuration/utils.js");
/* harmony import */ var _Presets__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./Presets */ "../tablero/lib/configuration/Presets.js");
/* harmony import */ var _Configuration__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./Configuration */ "../tablero/lib/configuration/Configuration.js");




const taskColumns = [
    { title: '', id: 'initialColumn-0', width: 40, type: _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].Checkbox },
    { title: 'Task title', id: 'initialColumn-1', width: 262, type: _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].SimpleRichText },
    { title: 'Assigned to', id: 'initialColumn-2', width: 190, type: _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].Text },
    { title: 'Due date', id: 'initialColumn-3', width: 190, type: _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].Date },
    { title: 'Planner Task Id', id: 'initialColumn-4', width: 10, type: _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].Text, isHidden: true },
    { title: 'Notes', id: 'initialColumn-5', width: 190, type: _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].RichText }
];
const taskRows = [Object(_utils__WEBPACK_IMPORTED_MODULE_1__["getTableRow"])('initialRow-0', taskColumns)];
const defaultNumbersRows = taskRows.length;
const presetOverrides = new Map([
    [_Presets__WEBPACK_IMPORTED_MODULE_2__["PresetName"].DisableInsertColumns, true],
    [_Presets__WEBPACK_IMPORTED_MODULE_2__["PresetName"].DisableDeleteColumns, true],
    [_Presets__WEBPACK_IMPORTED_MODULE_2__["PresetName"].DisableRenameColumns, true],
    [_Presets__WEBPACK_IMPORTED_MODULE_2__["PresetName"].EnableAddRowButton, true]
]);
const tasks = {
    configurationType: _Configuration__WEBPACK_IMPORTED_MODULE_3__["Configuration"].tasks,
    createType: 'newTablero',
    componentFriendlyName: 'Action Items',
    columns: taskColumns,
    rows: taskRows
};
const tasksPresets = Object(_Presets__WEBPACK_IMPORTED_MODULE_2__["overrideDefaultPresets"])(presetOverrides);


/***/ }),

/***/ "../tablero/lib/configuration/utils.js":
/*!*********************************************!*\
  !*** ../tablero/lib/configuration/utils.js ***!
  \*********************************************/
/*! exports provided: getTableRow, getTableroColumns, getTableroRows */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getTableRow", function() { return getTableRow; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getTableroColumns", function() { return getTableroColumns; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getTableroRows", function() { return getTableroRows; });
/* harmony import */ var _dataModel_adapterComponents_TableroAdapterComponentFactory__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../dataModel/adapterComponents/TableroAdapterComponentFactory */ "../tablero/lib/dataModel/adapterComponents/TableroAdapterComponentFactory.js");

const columnIdPrefix = 'initialColumn-';
const rowIdPrefix = 'initialRow-';
/**
 * Helper utility for generating a table row given table columns
 * @param id Row identifier
 * @param columns table column specifications
 */
const getTableRow = (id, columns) => {
    let initialData = {};
    columns.forEach((column) => {
        initialData[column.id] = Object(_dataModel_adapterComponents_TableroAdapterComponentFactory__WEBPACK_IMPORTED_MODULE_0__["createAdapterComponent"])(column.type);
    });
    return {
        id,
        data: initialData
    };
};
const getTableroColumns = (columnTypes, columnWidth) => {
    const cols = new Array(columnTypes.length);
    for (let colIndex = 0; colIndex < columnTypes.length; colIndex += 1) {
        cols[colIndex] = {
            title: '',
            id: columnIdPrefix + colIndex,
            width: columnWidth,
            type: columnTypes[colIndex]
        };
    }
    return cols;
};
const getTableroRows = (nRows, columns) => {
    const rows = new Array(nRows);
    for (let rowIndex = 0; rowIndex < nRows; rowIndex += 1) {
        rows[rowIndex] = getTableRow(rowIdPrefix + rowIndex, columns);
    }
    return rows;
};


/***/ }),

/***/ "../tablero/lib/dataModel/CellDataUtils.js":
/*!*************************************************!*\
  !*** ../tablero/lib/dataModel/CellDataUtils.js ***!
  \*************************************************/
/*! exports provided: createDefaultCellData, getCellDisplayText, useCellDisplayText, getCellValue, isEmptyCell, isNaNCell, tryGetComponentUrl */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "createDefaultCellData", function() { return createDefaultCellData; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getCellDisplayText", function() { return getCellDisplayText; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "useCellDisplayText", function() { return useCellDisplayText; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getCellValue", function() { return getCellValue; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "isEmptyCell", function() { return isEmptyCell; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "isNaNCell", function() { return isNaNCell; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tryGetComponentUrl", function() { return tryGetComponentUrl; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _view_table_Table_props__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../view/table/Table.props */ "../tablero/lib/view/table/Table.props.js");
/* harmony import */ var _view_StringTable__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../view/StringTable */ "../tablero/lib/view/StringTable.js");
/* harmony import */ var _adapterComponents_TableroAdapterComponentFactory__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./adapterComponents/TableroAdapterComponentFactory */ "../tablero/lib/dataModel/adapterComponents/TableroAdapterComponentFactory.js");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../telemetry */ "../tablero/lib/telemetry/ErrorEvents.js");
/* harmony import */ var _fluidx_utilities__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! @fluidx/utilities */ "../utilities/lib/logger/LoggerHelpers.js");






/**
 * Create new cell data object from a dataType
 * @param dataType
 */
const createDefaultCellData = (dataType) => {
    return Object(_adapterComponents_TableroAdapterComponentFactory__WEBPACK_IMPORTED_MODULE_3__["createAdapterComponent"])(dataType);
};
const getCellDisplayText = async (cellData, columnType = _view_table_Table_props__WEBPACK_IMPORTED_MODULE_1__["DataType"].Text) => {
    const cellValue = await getCellValue(cellData, columnType);
    if (cellValue) {
        switch (columnType) {
            case _view_table_Table_props__WEBPACK_IMPORTED_MODULE_1__["DataType"].Text: {
                return cellValue.toString();
            }
            case _view_table_Table_props__WEBPACK_IMPORTED_MODULE_1__["DataType"].Checkbox: {
                return cellValue ? _view_StringTable__WEBPACK_IMPORTED_MODULE_2__["checkedDisplayText"] : _view_StringTable__WEBPACK_IMPORTED_MODULE_2__["uncheckedDisplayText"];
            }
        }
    }
    return '';
};
/**
 * Hook wrapper around getCellDisplayText for react consumers of cell value as text
 */
const useCellDisplayText = (cellData) => {
    const [displayText, setDisplayText] = react__WEBPACK_IMPORTED_MODULE_0__["useState"]('');
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        if (cellData) {
            // tslint:disable-next-line: no-floating-promises
            getCellDisplayText(cellData, _view_table_Table_props__WEBPACK_IMPORTED_MODULE_1__["DataType"].Text).then((value) => setDisplayText(value));
        }
    }, [cellData]);
    return displayText;
};
const getCellValue = async (cellData, columnType, logger) => {
    switch (columnType) {
        case _view_table_Table_props__WEBPACK_IMPORTED_MODULE_1__["DataType"].Text: {
            const stringProducer = cellData.ComponentProduceString;
            return stringProducer !== undefined ? stringProducer.getString() : (logInvalidProducerNesting(logger), undefined);
        }
        case _view_table_Table_props__WEBPACK_IMPORTED_MODULE_1__["DataType"].Checkbox: {
            const booleanProducer = cellData.ComponentProduceBoolean;
            return booleanProducer !== undefined
                ? booleanProducer.getBoolean()
                : (logInvalidProducerNesting(logger), undefined);
        }
        case _view_table_Table_props__WEBPACK_IMPORTED_MODULE_1__["DataType"].Date: {
            const dateProducer = cellData.ComponentProduceDate;
            return dateProducer !== undefined ? dateProducer.getDate() : (logInvalidProducerNesting(logger), undefined);
        }
        case _view_table_Table_props__WEBPACK_IMPORTED_MODULE_1__["DataType"].RichText:
        case _view_table_Table_props__WEBPACK_IMPORTED_MODULE_1__["DataType"].SimpleRichText: {
            const richTextEditor = cellData.ComponentRichTextEditor;
            return richTextEditor !== undefined ? richTextEditor.getText() : (logInvalidProducerNesting(logger), undefined);
        }
    }
};
const logInvalidProducerNesting = (logger) => {
    logger &&
        Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_5__["sendErrorEvent"])(logger, {
            eventName: _telemetry__WEBPACK_IMPORTED_MODULE_4__["ErrorEvent"].ComponentNestingError,
            errorCode: _telemetry__WEBPACK_IMPORTED_MODULE_4__["ErrorCode"].ComponentNestingError.NestIsNotAValidProducer
        });
};
/**
 * Checks if a given cell has value of undefined or empty string
 * @param cellData CellData
 */
const isEmptyCell = async (cellData, columnType) => {
    const cellValue = await getCellValue(cellData, columnType);
    return cellValue === undefined || cellValue === '';
};
/**
 * Checks if a given cell has a non-numerical value
 * @param cellData CellData
 */
const isNaNCell = async (cellData, columnType) => {
    const cellValue = await getCellValue(cellData, columnType);
    return isNaN(Number(cellValue));
};
/**
 * Try and get url if cell data component is loadable
 * @param cellData CellData
 */
const tryGetComponentUrl = (cellData) => {
    const component = cellData;
    if (!component || !component.IComponentLoadable) {
        return;
    }
    return component.IComponentLoadable.url;
};


/***/ }),

/***/ "../tablero/lib/dataModel/FluidConstants.js":
/*!**************************************************!*\
  !*** ../tablero/lib/dataModel/FluidConstants.js ***!
  \**************************************************/
/*! exports provided: tableDocumentIdKey, tableroDocumentIdKey */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableDocumentIdKey", function() { return tableDocumentIdKey; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableroDocumentIdKey", function() { return tableroDocumentIdKey; });
const tableDocumentIdKey = 'tableDocumentId';
const tableroDocumentIdKey = 'tableroDocumentId';


/***/ }),

/***/ "../tablero/lib/dataModel/Properties.js":
/*!**********************************************!*\
  !*** ../tablero/lib/dataModel/Properties.js ***!
  \**********************************************/
/*! exports provided: Properties */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Properties", function() { return Properties; });
const Properties = {
    id: 'id',
    type: 'type',
    title: 'title'
};


/***/ }),

/***/ "../tablero/lib/dataModel/TableSorting.js":
/*!************************************************!*\
  !*** ../tablero/lib/dataModel/TableSorting.js ***!
  \************************************************/
/*! exports provided: sortRows, stripRowComparisonValues */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "sortRows", function() { return sortRows; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "stripRowComparisonValues", function() { return stripRowComparisonValues; });
/* harmony import */ var _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../view/table/Table.props */ "../tablero/lib/view/table/Table.props.js");
/* harmony import */ var _CellDataUtils__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./CellDataUtils */ "../tablero/lib/dataModel/CellDataUtils.js");


/**
 * For performance reasons, we use a single collator object for comparisons instead of using localeCompare
 * More info: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/String/localeCompare#Performance
 */
const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });
/**
 * Function that can be used in conjunction with Array.prototype.sort to sort cells
 */
const tableCellComparer = (first, second, columnType, sortDirection) => {
    // The undefined means that the component from that cell is not producing the expected type.
    // In that case, those values will be shown in the last rows.
    if (first === undefined && second === undefined) {
        return 0;
    }
    if (first === undefined) {
        return 1;
    }
    if (second === undefined) {
        return -1;
    }
    let compareValue = 0;
    switch (columnType) {
        case _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].Text:
            compareValue = collator.compare(first, second);
            break;
        case _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].Checkbox:
            const firstAsNumber = first ? 1 : 0;
            const secondAsNumber = second ? 1 : 0;
            compareValue = firstAsNumber - secondAsNumber;
            break;
        case _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].Date:
            compareValue = first.getTime() - second.getTime();
            break;
        case _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].RichText:
        case _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].SimpleRichText:
            compareValue = collator.compare(first, second);
            break;
    }
    return sortDirection === _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["SortDirection"].Descending ? -1 * compareValue : compareValue;
};
/**
 * Performs a sort on a dataCell based array in three steps:
 * 1. Schwartzian transform the data array by asynchronously pre-fetching sortable data
 * to be used as comparison values from cellData in array elements
 *
 * 2. Sort the transformed data array: `comparableData: (CellValueType | Row)[][]`
 * where `comparableData[0] : CellValueType` is the comparison value
 * and `comparableData[1]: T` is the original dataArray elements to be sorted
 *
 * 3. Strip comparison values and return a sorted array
 *
 * @param rows `Row[]` Array to be sorted based on cellData
 * @param columnType DataType of column on which data array is to be sorted
 * @param sortInformation : Additional context information used when sorting complex data array like Row[] based on nested cellData
 */
const sortRows = async (rows, columnType, sortInformation) => {
    const comparableRows = await Promise.all(rows.map(async (row) => {
        const cellValue = await Object(_CellDataUtils__WEBPACK_IMPORTED_MODULE_1__["getCellValue"])(row.data[sortInformation.sortByColumnId], columnType);
        return { sortKey: cellValue, element: row };
    }));
    // At this step, the comparableRows array contains only elements which either have the desired type based on the column type or are undefined.
    // For example, if we are in a Date column we are expecting to have only Dates.
    // In case other components land there, if they don't produce Date type, their corresponding value from comparableRows will be undefined.
    // This allows us to use tableCellComparer without checking the typeof produced value, because we already made sure to have the same type.
    const sortedRows = comparableRows.sort((first, second) => tableCellComparer(first.sortKey, second.sortKey, columnType, sortInformation.sortDirection));
    return stripRowComparisonValues(sortedRows);
};
/**
 * Strips the comparison values from an already sorted array.
 * @param sortedRows The rows array from a Schwartzian transform.
 *
 * @returns An array constructed from the corresponding elements.
 */
const stripRowComparisonValues = (sortedRows) => {
    return sortedRows.map((rowData) => rowData.element);
};
// cSpell:ignore Schwartzian


/***/ }),

/***/ "../tablero/lib/dataModel/TableroAdapter.js":
/*!**************************************************!*\
  !*** ../tablero/lib/dataModel/TableroAdapter.js ***!
  \**************************************************/
/*! exports provided: nextRowIdKey, componentConfigurationTypeKey, getCellData, setCellValue, getColumnDataProperties, getColumnData, getRowData, getColumnViewProperties, setColumnViewProperties, getRowViewProperties, setRowViewProperties, getSortByColumn, setSortByColumn, getHiddenColumnIds, setHiddenColumnIds, generateRowId, generateTableroDocumentNextRowId, generateColumnId, insertNewRow, insertNewColumn, deleteRow, deleteColumn, cacheIdsToIndex, getOrderedRowIndexFromId, updateColumnData, updateColumnViewProperties, getColumnIfVisible, getHeaderDataForRender, getBodyDataForRender, getColumnTitle, getCellsInColumn, getColumnType, getNumericCellsInColumn, getTextCellsInColumn, getPossibleFilterValuesForColumn, setComponentFriendlyName, getComponentFriendlyName, getTableroDocumentId, setTableroDocumentId, setComponentConfigurationType, getComponentConfigurationType, addHiddenColumn, getAttributionUserFromCellProperties */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "nextRowIdKey", function() { return nextRowIdKey; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "componentConfigurationTypeKey", function() { return componentConfigurationTypeKey; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getCellData", function() { return getCellData; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "setCellValue", function() { return setCellValue; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getColumnDataProperties", function() { return getColumnDataProperties; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getColumnData", function() { return getColumnData; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getRowData", function() { return getRowData; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getColumnViewProperties", function() { return getColumnViewProperties; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "setColumnViewProperties", function() { return setColumnViewProperties; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getRowViewProperties", function() { return getRowViewProperties; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "setRowViewProperties", function() { return setRowViewProperties; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getSortByColumn", function() { return getSortByColumn; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "setSortByColumn", function() { return setSortByColumn; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getHiddenColumnIds", function() { return getHiddenColumnIds; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "setHiddenColumnIds", function() { return setHiddenColumnIds; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "generateRowId", function() { return generateRowId; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "generateTableroDocumentNextRowId", function() { return generateTableroDocumentNextRowId; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "generateColumnId", function() { return generateColumnId; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "insertNewRow", function() { return insertNewRow; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "insertNewColumn", function() { return insertNewColumn; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "deleteRow", function() { return deleteRow; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "deleteColumn", function() { return deleteColumn; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "cacheIdsToIndex", function() { return cacheIdsToIndex; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getOrderedRowIndexFromId", function() { return getOrderedRowIndexFromId; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "updateColumnData", function() { return updateColumnData; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "updateColumnViewProperties", function() { return updateColumnViewProperties; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getColumnIfVisible", function() { return getColumnIfVisible; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getHeaderDataForRender", function() { return getHeaderDataForRender; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getBodyDataForRender", function() { return getBodyDataForRender; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getColumnTitle", function() { return getColumnTitle; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getCellsInColumn", function() { return getCellsInColumn; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getColumnType", function() { return getColumnType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getNumericCellsInColumn", function() { return getNumericCellsInColumn; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getTextCellsInColumn", function() { return getTextCellsInColumn; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getPossibleFilterValuesForColumn", function() { return getPossibleFilterValuesForColumn; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "setComponentFriendlyName", function() { return setComponentFriendlyName; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getComponentFriendlyName", function() { return getComponentFriendlyName; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getTableroDocumentId", function() { return getTableroDocumentId; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "setTableroDocumentId", function() { return setTableroDocumentId; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "setComponentConfigurationType", function() { return setComponentConfigurationType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getComponentConfigurationType", function() { return getComponentConfigurationType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "addHiddenColumn", function() { return addHiddenColumn; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getAttributionUserFromCellProperties", function() { return getAttributionUserFromCellProperties; });
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/array.js");
/* harmony import */ var _fluidx_utilities__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @fluidx/utilities */ "../utilities/lib/logger/LoggerHelpers.js");
/* harmony import */ var _fluidx_utilities__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @fluidx/utilities */ "../utilities/lib/attribution/DefaultAttributionProvider.js");
/* harmony import */ var _TableSorting__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./TableSorting */ "../tablero/lib/dataModel/TableSorting.js");
/* harmony import */ var _view_rows_RowFiltering__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../view/rows/RowFiltering */ "../tablero/lib/view/rows/RowFiltering.js");
/* harmony import */ var _Properties__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ./Properties */ "../tablero/lib/dataModel/Properties.js");
/* harmony import */ var _FluidConstants__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ./FluidConstants */ "../tablero/lib/dataModel/FluidConstants.js");
/* harmony import */ var _view_table_Table_props__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../view/table/Table.props */ "../tablero/lib/view/table/Table.props.js");
/* harmony import */ var _adapterComponents_TableroAdapterComponentFactory__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ./adapterComponents/TableroAdapterComponentFactory */ "../tablero/lib/dataModel/adapterComponents/TableroAdapterComponentFactory.js");
/* harmony import */ var _CellDataUtils__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ./CellDataUtils */ "../tablero/lib/dataModel/CellDataUtils.js");
/* harmony import */ var _configuration_Presets__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ../configuration/Presets */ "../tablero/lib/configuration/Presets.js");
/* harmony import */ var _configuration_Configuration__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(/*! ../configuration/Configuration */ "../tablero/lib/configuration/Configuration.js");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_12__ = __webpack_require__(/*! ../telemetry */ "../tablero/lib/telemetry/ErrorEvents.js");
/* harmony import */ var _TableroDocumentUtil__WEBPACK_IMPORTED_MODULE_13__ = __webpack_require__(/*! ./TableroDocumentUtil */ "../tablero/lib/dataModel/TableroDocumentUtil.js");













// TableroView Keys
const orderedRowIdsKey = 'orderedRowIds';
const nextRowIdKey = 'nextRowId';
const nextColumnIdKey = 'nextColumnId';
const orderedColumnsKey = 'orderedColumns';
const sortedColumnKey = 'sortedColumn';
const hiddenColumnIdsKey = 'hiddenColumnIds';
const componentFriendlyNameMapKey = 'componentFriendlyName';
const componentConfigurationTypeKey = 'componentConfigurationType';
const getCellData = async (tableroDocument, rowId, columnId, columnType, context, logger) => {
    let rowIndex = Object(_TableroDocumentUtil__WEBPACK_IMPORTED_MODULE_13__["getIndexFromId"])(tableroDocument.rowIdToIndex, rowId);
    let colIndex = Object(_TableroDocumentUtil__WEBPACK_IMPORTED_MODULE_13__["getIndexFromId"])(tableroDocument.colIdToIndex, columnId);
    const table = tableroDocument.table;
    const tableCellValue = table.getCellValue(rowIndex, colIndex);
    let componentUrl;
    let cellValue;
    if (tableCellValue !== undefined) {
        try {
            const parsedCellData = JSON.parse(tableCellValue);
            componentUrl = parsedCellData.componentUrl;
            cellValue = parsedCellData.cellValue;
        }
        catch (error) {
            logger &&
                Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_1__["sendErrorEvent"])(logger, {
                    eventName: _telemetry__WEBPACK_IMPORTED_MODULE_12__["ErrorEvent"].DataInitializationError,
                    errorCode: _telemetry__WEBPACK_IMPORTED_MODULE_12__["ErrorCode"].DataInitializationError.ErrorParsingCellValue
                }, error);
        }
    }
    if (componentUrl) {
        const response = await context.containerRuntime.request({ url: componentUrl });
        if (response.status !== 200 || response.mimeType !== 'fluid/component') {
            // In case of negative response status, instead of returning undefined CellData
            // We are returning a default Adapter Component corresponding to the column type to handle Nested Component N/A issues
            // Logging the ComponentNestingError also.
            logger &&
                Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_1__["sendErrorEvent"])(logger, {
                    eventName: _telemetry__WEBPACK_IMPORTED_MODULE_12__["ErrorEvent"].ComponentNestingError,
                    errorCode: _telemetry__WEBPACK_IMPORTED_MODULE_12__["ErrorCode"].ComponentNestingError.NestFailedToLoad
                });
        }
        else {
            return response.value;
        }
    }
    return Object(_adapterComponents_TableroAdapterComponentFactory__WEBPACK_IMPORTED_MODULE_8__["createAdapterComponent"])(columnType, cellValue);
};
const serializeCellValue = async (cellData, columnType, logger) => {
    const cellValue = await Object(_CellDataUtils__WEBPACK_IMPORTED_MODULE_9__["getCellValue"])(cellData, columnType, logger);
    const componentUrl = Object(_CellDataUtils__WEBPACK_IMPORTED_MODULE_9__["tryGetComponentUrl"])(cellData);
    return JSON.stringify({ componentUrl, cellValue });
};
const setCellValue = async (tableroDocument, rowId, columnId, value, logger) => {
    let rowIndex = Object(_TableroDocumentUtil__WEBPACK_IMPORTED_MODULE_13__["getIndexFromId"])(tableroDocument.rowIdToIndex, rowId);
    let colIndex = Object(_TableroDocumentUtil__WEBPACK_IMPORTED_MODULE_13__["getIndexFromId"])(tableroDocument.colIdToIndex, columnId);
    let columnType = getColumnType(tableroDocument, columnId);
    const cellStringValue = await serializeCellValue(value, columnType, logger);
    if (tableroDocument.tableWithInterception) {
        // Add user details to the fluid table
        tableroDocument.tableWithInterception.setCellValue(rowIndex, colIndex, cellStringValue);
    }
    else {
        tableroDocument.table.setCellValue(rowIndex, colIndex, cellStringValue);
    }
};
const getColumnDataProperties = (tableroDocument, id) => {
    let properties = tableroDocument.table.getColProperties(Object(_TableroDocumentUtil__WEBPACK_IMPORTED_MODULE_13__["getIndexFromId"])(tableroDocument.colIdToIndex, id));
    if (properties === undefined) {
        throw new Error('Attempting to access column data of a column with no properties.');
    }
    return { id: properties[_Properties__WEBPACK_IMPORTED_MODULE_5__["Properties"].id], type: properties[_Properties__WEBPACK_IMPORTED_MODULE_5__["Properties"].type], title: properties[_Properties__WEBPACK_IMPORTED_MODULE_5__["Properties"].title] };
};
const getColumnData = (tablero) => {
    const viewColumns = tablero.tableroViewStore.rootMap.get(orderedColumnsKey);
    return viewColumns.map((column) => {
        const tableroDocumentColumn = getColumnDataProperties(tablero.tableroDocumentStore, column.id);
        return {
            id: column.id,
            // Document Properties
            title: tableroDocumentColumn.title,
            type: tableroDocumentColumn.type,
            // View Properties
            filter: column.filter,
            width: column.width,
            sortDirection: column.sortDirection,
            isHidden: column.isHidden
        };
    });
};
const getRowData = async (tablero, columnData, context, logger) => {
    const orderedRowIds = getRowViewProperties(tablero.tableroViewStore);
    const orderedRows = orderedRowIds.map(async (row) => {
        const result = {
            id: row.id,
            data: {}
        };
        await Promise.all(columnData.map(async (column) => {
            result.data[column.id] = await getCellData(tablero.tableroDocumentStore, row.id, column.id, column.type, context, logger);
        }));
        return result;
    });
    return Promise.all(orderedRows);
};
const getColumnViewProperties = (tableroViewStore) => {
    return tableroViewStore.rootMap.get(orderedColumnsKey);
};
const setColumnViewProperties = (tableroViewStore, columns) => {
    tableroViewStore.rootMap.set(orderedColumnsKey, columns);
};
const getRowViewProperties = (tableroViewStore) => {
    return tableroViewStore.rootMap.get(orderedRowIdsKey);
};
function setRowViewProperties(tableroViewStore, rowIds) {
    tableroViewStore.rootMap.set(orderedRowIdsKey, rowIds);
}
const getSortByColumn = (tableroViewStore) => tableroViewStore.rootMap.get(sortedColumnKey);
function setSortByColumn(tablero, sortInformation) {
    tablero.rootMap.set(sortedColumnKey, sortInformation);
}
const getHiddenColumnIds = (tablero) => tablero.rootMap.get(hiddenColumnIdsKey) || [];
function setHiddenColumnIds(tablero, hiddenColumnIds) {
    tablero.rootMap.set(hiddenColumnIdsKey, hiddenColumnIds);
}
const getNextIdNumber = (tableroDocument, nextIdKey) => {
    const nextIdNum = tableroDocument.rootMap.get(nextIdKey) || 0;
    tableroDocument.rootMap.set(nextIdKey, nextIdNum + 1);
    return nextIdNum;
};
// Generates a new row id to use for a new row that should not collide with other ids
const generateRowId = (tableroDocument) => {
    const nextIdNum = getNextIdNumber(tableroDocument, nextRowIdKey);
    return `row-${nextIdNum}`;
};
const generateTableroDocumentNextRowId = (tableroDocumentStore) => {
    // Traverse over all rowId and find next highest possible rowId
    let rowId = '';
    let nextRowID;
    let previousRowId;
    let numericRowId = 0;
    let newRowAdded = false;
    for (let rowIndex in tableroDocumentStore.rowIdToIndex) {
        if (rowIndex !== undefined) {
            rowId = rowIndex;
            if (rowId.startsWith('row')) {
                newRowAdded = true;
            }
            previousRowId = parseInt(rowId === null || rowId === void 0 ? void 0 : rowId.substr((rowId === null || rowId === void 0 ? void 0 : rowId.indexOf('-')) + 1), 10);
            numericRowId = numericRowId > previousRowId ? numericRowId : previousRowId;
        }
    }
    if (!newRowAdded) {
        nextRowID = 0;
    }
    else {
        nextRowID = numericRowId + 1;
    }
    return nextRowID;
};
// Generates a new id that can be used to create a new column that should not collide with other ids.
const generateColumnId = (tableroDocument) => {
    const nextIdNum = getNextIdNumber(tableroDocument, nextColumnIdKey);
    return `column-${nextIdNum}`;
};
function insertNewRow(tableroDocumentStore, newRowId, index) {
    tableroDocumentStore.table.insertRows(index, 1);
    // Annotate new row with id
    let propertyBag = {
        [_Properties__WEBPACK_IMPORTED_MODULE_5__["Properties"].id]: newRowId
    };
    tableroDocumentStore.table.annotateRows(index, index + 1, propertyBag);
    cacheIdsToIndex(tableroDocumentStore, true /*isRows*/);
}
function insertNewColumn(tableroDocumentStore, newColId, index, columnDataType) {
    tableroDocumentStore.table.insertCols(index, 1);
    // Annotate new columns
    let propertyBag = {
        [_Properties__WEBPACK_IMPORTED_MODULE_5__["Properties"].id]: newColId,
        [_Properties__WEBPACK_IMPORTED_MODULE_5__["Properties"].type]: columnDataType,
        [_Properties__WEBPACK_IMPORTED_MODULE_5__["Properties"].title]: ''
    };
    tableroDocumentStore.table.annotateCols(index, index + 1, propertyBag);
    cacheIdsToIndex(tableroDocumentStore, false /*isRows*/);
}
function deleteRow(tableroDocumentStore, index) {
    tableroDocumentStore.table.removeRows(index, 1 /*numRows*/);
    cacheIdsToIndex(tableroDocumentStore, true /*isRows*/);
}
function deleteColumn(tableroDocumentStore, index) {
    tableroDocumentStore.table.removeCols(index, 1 /*numCols*/);
    cacheIdsToIndex(tableroDocumentStore, false /*isRows*/);
}
function cacheIdsToIndex(documentStore, isRows) {
    const idToIndexMap = {};
    const table = documentStore.table;
    const numRows = table.numRows;
    const numCols = table.numCols;
    const maxRow = numRows - 1;
    const maxCol = numCols - 1;
    const startVal = 0;
    const endVal = isRows ? maxRow : maxCol;
    for (let start = startVal; start < endVal; start += 1) {
        const properties = isRows ? table.getRowProperties(start) : table.getColProperties(start);
        let id = '';
        if (properties !== undefined) {
            // Right now we assume if it has properties it has the ID property - true for now, but may not always be true.
            id = properties[_Properties__WEBPACK_IMPORTED_MODULE_5__["Properties"].id];
            idToIndexMap[id] = start;
        }
    }
    if (isRows) {
        documentStore.rowIdToIndex = idToIndexMap;
    }
    else {
        documentStore.colIdToIndex = idToIndexMap;
    }
}
const getOrderedRowIndexFromId = (tableroViewStore, id) => {
    const orderedRowIds = getRowViewProperties(tableroViewStore);
    for (let index = 0; index < orderedRowIds.length; index += 1) {
        if (orderedRowIds[index].id === id)
            return index;
    }
    throw new Error('We expect that id exists.');
};
const updateColumnData = (documentStore, columnId, columnOverrides) => {
    const column = getColumnDataProperties(documentStore, columnId);
    let newColumnData = Object.assign(Object.assign({}, column), columnOverrides);
    let columnIndex = Object(_TableroDocumentUtil__WEBPACK_IMPORTED_MODULE_13__["getIndexFromId"])(documentStore.colIdToIndex, columnId);
    let propertyBag = {
        [_Properties__WEBPACK_IMPORTED_MODULE_5__["Properties"].id]: newColumnData.id,
        [_Properties__WEBPACK_IMPORTED_MODULE_5__["Properties"].type]: newColumnData.type,
        [_Properties__WEBPACK_IMPORTED_MODULE_5__["Properties"].title]: newColumnData.title
    };
    documentStore.table.annotateCols(columnIndex, columnIndex + 1, propertyBag);
};
const updateColumnViewProperties = (tableroViewStore, columnId, columnOverrides) => {
    const columns = getColumnViewProperties(tableroViewStore);
    const columnIndex = columns.findIndex((c) => c.id === columnId);
    if (columnIndex >= 0) {
        const newColumns = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["replaceElement"])(columns, Object.assign(Object.assign({}, columns[columnIndex]), columnOverrides), columnIndex);
        setColumnViewProperties(tableroViewStore, newColumns);
    }
};
const getVisibleColumns = (columns, tableroViewStore) => {
    const hiddenColumnIds = new Set(getHiddenColumnIds(tableroViewStore));
    return columns.filter((c) => !hiddenColumnIds.has(c.id));
};
const tableHasHiddenColumns = (viewStore) => {
    return getHiddenColumnIds(viewStore).length > 0;
};
function getColumnIfVisible(tablero, id) {
    const columns = getColumnData(tablero);
    const visibleColumns = getVisibleColumns(columns, tablero.tableroViewStore);
    return visibleColumns.find((x) => x.id === id);
}
const getHeaderDataForRender = (tablero) => {
    refreshIdCacheIfNecessary(tablero);
    let orderedColumns = getColumnData(tablero);
    const sortInformation = getSortByColumn(tablero.tableroViewStore);
    if (sortInformation && sortInformation.sortDirection) {
        const colIndex = orderedColumns.findIndex((c) => c.id === sortInformation.sortByColumnId);
        orderedColumns = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["replaceElement"])(orderedColumns, Object.assign(Object.assign({}, orderedColumns[colIndex]), { sortDirection: sortInformation.sortDirection }), colIndex);
    }
    const visibleColumns = getVisibleColumns(orderedColumns, tablero.tableroViewStore);
    const hasHiddenColumns = tableHasHiddenColumns(tablero.tableroViewStore);
    // TODO: TABLERO: Handle the null component title case.
    const title = getComponentFriendlyName(tablero.tableroViewStore) || '';
    return { columns: visibleColumns, title, hasHiddenColumns };
};
const getBodyDataForRender = async (tablero, context, logger) => {
    refreshIdCacheIfNecessary(tablero);
    let orderedColumns = getColumnData(tablero);
    let orderedRows = await getRowData(tablero, orderedColumns, context, logger);
    let filteredRows = await Object(_view_rows_RowFiltering__WEBPACK_IMPORTED_MODULE_4__["filterRows"])(orderedRows, orderedColumns);
    const sortInformation = getSortByColumn(tablero.tableroViewStore);
    if (sortInformation && sortInformation.applySort && sortInformation.sortDirection) {
        const columnType = getColumnType(tablero.tableroDocumentStore, sortInformation.sortByColumnId);
        // Sort filtered rows
        filteredRows = (await Object(_TableSorting__WEBPACK_IMPORTED_MODULE_3__["sortRows"])(filteredRows, columnType, sortInformation));
        // Updating Row View Properties after Sort is applied since we are not rearranging
        // on every re-render and thus saving the sorted state.
        // TODO: Revisit Sort functionality with broader fix in Tablero Redesign.
        let rowViewProperties = [];
        filteredRows.forEach((row) => {
            rowViewProperties.push({ id: row.id });
        });
        setRowViewProperties(tablero.tableroViewStore, rowViewProperties);
        // Disable applySort flag since we are not rearranging on every re-render.
        // Saving sortDirection and sortByColumnId to allow toggle button to work with current implementation.
        // TODO: Fix this properly when Toggle button User Behavior is thought through for Sorting.
        setSortByColumn(tablero.tableroViewStore, {
            sortByColumnId: sortInformation.sortByColumnId,
            sortDirection: sortInformation.sortDirection,
            applySort: false
        });
    }
    return { rows: filteredRows };
};
/**
 * We maintain a mapping of IDs to TableDocument indexes in order to refer to constant IDs in our view logic but
 * be able to use TableDocument methods. The cache needs to be refreshed whenever a structural row/col change occurs.
 */
const refreshIdCacheIfNecessary = (tableroStore) => {
    const { tableroDocumentStore, tableroViewStore } = tableroStore;
    let knownRowIds = Object.keys(tableroDocumentStore.rowIdToIndex);
    let knownColIds = Object.keys(tableroDocumentStore.colIdToIndex);
    const currentRowIds = getRowViewProperties(tableroViewStore).map((row) => row.id);
    const currentColumnIds = getColumnViewProperties(tableroViewStore).map((column) => column.id);
    // Check to see if the lengths match and if the values are the same. If not, recache.
    if (!(currentRowIds.length === knownRowIds.length &&
        currentRowIds.every(function (rowId) {
            return knownRowIds.includes(rowId);
        }, knownRowIds))) {
        cacheIdsToIndex(tableroDocumentStore, true /* isRows */);
    }
    if (!(currentColumnIds.length === knownColIds.length &&
        currentColumnIds.every(function (columnId) {
            return knownColIds.includes(columnId);
        }, knownColIds))) {
        cacheIdsToIndex(tableroDocumentStore, false /* isRows */);
    }
};
const getColumnTitle = (tableroDocumentStore, columnId) => {
    return getColumnDataProperties(tableroDocumentStore, columnId).title;
};
const getCellsInColumn = async (tablero, columnId, context) => {
    const rowViewProperties = getRowViewProperties(tablero.tableroViewStore);
    let columnType = getColumnType(tablero.tableroDocumentStore, columnId);
    const cellData = rowViewProperties.map(async (row) => await getCellData(tablero.tableroDocumentStore, row.id, columnId, columnType, context));
    return Promise.all(cellData);
};
function getColumnType(tableroDocumentStore, columnId) {
    const column = getColumnDataProperties(tableroDocumentStore, columnId);
    return column.type;
}
const getNumericCellsInColumn = async (tablero, columnId, context) => {
    const columnType = getColumnType(tablero.tableroDocumentStore, columnId);
    // Assume all components in a column must produce the same type
    if (columnType !== _view_table_Table_props__WEBPACK_IMPORTED_MODULE_7__["DataType"].Text) {
        return [];
    }
    const cellData = await getCellsInColumn(tablero, columnId, context);
    const numericCellValues = [];
    await Promise.all(cellData.map(async (cell) => {
        if (!(await Object(_CellDataUtils__WEBPACK_IMPORTED_MODULE_9__["isNaNCell"])(cell, columnType))) {
            const cellValue = await Object(_CellDataUtils__WEBPACK_IMPORTED_MODULE_9__["getCellValue"])(cell, columnType);
            return numericCellValues.push(Number(cellValue));
        }
        return undefined;
    }));
    return numericCellValues;
};
const getTextCellsInColumn = async (tablero, columnId, context) => {
    const columnType = getColumnType(tablero.tableroDocumentStore, columnId);
    // Assume all components in a column must produce the same type
    if (columnType !== _view_table_Table_props__WEBPACK_IMPORTED_MODULE_7__["DataType"].Text) {
        return [];
    }
    const cellData = await getCellsInColumn(tablero, columnId, context);
    return Promise.all(cellData.map(async (cell) => (await Object(_CellDataUtils__WEBPACK_IMPORTED_MODULE_9__["getCellValue"])(cell, columnType))));
};
const getPossibleFilterValuesForColumn = async (tablero, columnId, context) => {
    const cellData = await getCellsInColumn(tablero, columnId, context);
    return Object(_view_rows_RowFiltering__WEBPACK_IMPORTED_MODULE_4__["getFilterSuggestionsForColumn"])(cellData);
};
const setComponentFriendlyName = (tableroViewStore, friendlyName) => {
    tableroViewStore.rootMap.set(componentFriendlyNameMapKey, friendlyName);
};
const getComponentFriendlyName = (tableroViewStore) => {
    return tableroViewStore.rootMap.get(componentFriendlyNameMapKey);
};
const getTableroDocumentId = (tableroViewStore) => {
    return tableroViewStore.rootMap.get(_FluidConstants__WEBPACK_IMPORTED_MODULE_6__["tableroDocumentIdKey"]);
};
function setTableroDocumentId(tableroViewStore, tableroDocumentId) {
    tableroViewStore.rootMap.set(_FluidConstants__WEBPACK_IMPORTED_MODULE_6__["tableroDocumentIdKey"], tableroDocumentId);
}
const setComponentConfigurationType = (tableroViewStore, configurationType) => {
    return tableroViewStore.rootMap.set(componentConfigurationTypeKey, configurationType);
};
const getComponentConfigurationType = (tableroViewStore, logger) => {
    let configurationType = tableroViewStore.rootMap.get(componentConfigurationTypeKey);
    //#region back compatibility
    // TODO: Task 3749578: Remove presets persistence back compatibility
    if (configurationType === undefined) {
        configurationType = approximateOldTableroConfigType(tableroViewStore, logger);
        setComponentConfigurationType(tableroViewStore, configurationType);
    }
    //#endregion back compatibility
    switch (configurationType) {
        case _configuration_Configuration__WEBPACK_IMPORTED_MODULE_11__["Configuration"].actionItemManager:
            return _configuration_Configuration__WEBPACK_IMPORTED_MODULE_11__["Configuration"].actionItemManager;
        case _configuration_Configuration__WEBPACK_IMPORTED_MODULE_11__["Configuration"].richTextTablero:
            return _configuration_Configuration__WEBPACK_IMPORTED_MODULE_11__["Configuration"].richTextTablero;
        case _configuration_Configuration__WEBPACK_IMPORTED_MODULE_11__["Configuration"].tasks:
            return _configuration_Configuration__WEBPACK_IMPORTED_MODULE_11__["Configuration"].tasks;
        default:
            return _configuration_Configuration__WEBPACK_IMPORTED_MODULE_11__["Configuration"].default;
    }
};
const addHiddenColumn = (tableroStore, columnId) => {
    const existingHiddenColumn = getHiddenColumnIds(tableroStore.tableroViewStore).filter((id) => id !== columnId);
    existingHiddenColumn.push(columnId);
    setHiddenColumnIds(tableroStore.tableroViewStore, existingHiddenColumn);
};
const getAttributionUserFromCellProperties = (tableroDocument, rowId, columnId) => {
    const rowIndex = Object(_TableroDocumentUtil__WEBPACK_IMPORTED_MODULE_13__["getIndexFromId"])(tableroDocument.rowIdToIndex, rowId);
    const colIndex = Object(_TableroDocumentUtil__WEBPACK_IMPORTED_MODULE_13__["getIndexFromId"])(tableroDocument.colIdToIndex, columnId);
    const properties = tableroDocument.table.getCellProperties(rowIndex, colIndex);
    if (properties && properties[_fluidx_utilities__WEBPACK_IMPORTED_MODULE_2__["propertyAttribution"]] !== undefined) {
        return properties[_fluidx_utilities__WEBPACK_IMPORTED_MODULE_2__["propertyAttribution"]];
    }
    return undefined;
};
//#region back compatibility
// TODO: Task 3749578: Remove presets persistence back compatibility
const tableroPresetNamesKey = 'orderedPresetNames';
const tableroPresetsKey = 'orderedPresets';
/**
 * We use the existence and value of `PresetName.EnableInsertColumns`, to guess the flavor of old tablero
 * Value of `false` points to this being a n action item manager and `true` to a default tablero
 * We could have also used `PresetName.EnableDeleteColumns`, (or both) to the same conclusion
 */
const approximateOldTableroConfigType = (tableroViewStore, logger) => {
    logger &&
        Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_1__["sendTelemetryEvent"])(logger, {
            eventName: _telemetry__WEBPACK_IMPORTED_MODULE_12__["ErrorEvent"].BackCompatibility,
            errorCode: _telemetry__WEBPACK_IMPORTED_MODULE_12__["ErrorCode"].BackCompatibility.SupportRootmapPresets
        });
    const rawPresetNames = tableroViewStore.rootMap.get(tableroPresetNamesKey);
    try {
        const presetNames = rawPresetNames && JSON.parse(rawPresetNames);
        const targetPresetIndex = presetNames && presetNames.indexOf(_configuration_Presets__WEBPACK_IMPORTED_MODULE_10__["PresetName"].EnableInsertColumns.toString());
        if (targetPresetIndex !== undefined && targetPresetIndex !== -1) {
            const presets = JSON.parse(tableroViewStore.rootMap.get(tableroPresetsKey));
            if (presets && targetPresetIndex < presets.length && !presets[targetPresetIndex]) {
                return _configuration_Configuration__WEBPACK_IMPORTED_MODULE_11__["Configuration"].actionItemManager;
            }
        }
    }
    catch (error) {
        logger &&
            Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_1__["sendErrorEvent"])(logger, { eventName: _telemetry__WEBPACK_IMPORTED_MODULE_12__["ErrorEvent"].Exception, errorCode: _telemetry__WEBPACK_IMPORTED_MODULE_12__["ErrorCode"].BackCompatibility.FailedToGetRootmapPresets }, error);
    }
    return _configuration_Configuration__WEBPACK_IMPORTED_MODULE_11__["Configuration"].default;
};
//#endregion back compatibility


/***/ }),

/***/ "../tablero/lib/dataModel/TableroDocumentUtil.js":
/*!*******************************************************!*\
  !*** ../tablero/lib/dataModel/TableroDocumentUtil.js ***!
  \*******************************************************/
/*! exports provided: insertNewTableRow, updateRowData, isValueComponentUrlType, updateCellData, deleteTableRow, getColumnData, getIndexFromId, getFluidUserWithDefaults */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "insertNewTableRow", function() { return insertNewTableRow; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "updateRowData", function() { return updateRowData; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "isValueComponentUrlType", function() { return isValueComponentUrlType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "updateCellData", function() { return updateCellData; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "deleteTableRow", function() { return deleteTableRow; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getColumnData", function() { return getColumnData; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getIndexFromId", function() { return getIndexFromId; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getFluidUserWithDefaults", function() { return getFluidUserWithDefaults; });
/* harmony import */ var _TableroAdapter__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./TableroAdapter */ "../tablero/lib/dataModel/TableroAdapter.js");
/* harmony import */ var util__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! util */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/util/0.11.1/node_modules/util/util.js");
/* harmony import */ var util__WEBPACK_IMPORTED_MODULE_1___default = /*#__PURE__*/__webpack_require__.n(util__WEBPACK_IMPORTED_MODULE_1__);


const insertNewTableRow = (tableroDocumentStore, data) => {
    // Assumation that new row will be added at the end only.
    let indexToInsert = tableroDocumentStore.table.numRows - 1;
    let nextRowID = Object(_TableroAdapter__WEBPACK_IMPORTED_MODULE_0__["generateTableroDocumentNextRowId"])(tableroDocumentStore);
    const newRowId = 'row-' + nextRowID;
    Object(_TableroAdapter__WEBPACK_IMPORTED_MODULE_0__["insertNewRow"])(tableroDocumentStore, newRowId, indexToInsert);
    // Populate data to new added row
    data.rowId = newRowId;
    updateRowData(tableroDocumentStore, data);
};
const updateRowData = (tableroDocumentStore, data) => {
    const rowIndex = getIndexFromId(tableroDocumentStore.rowIdToIndex, data.rowId);
    const columnData = getColumnData(tableroDocumentStore);
    columnData.forEach(async (column) => {
        let newCellValue = data.rowData[column];
        let columnIndex = getIndexFromId(tableroDocumentStore.colIdToIndex, column);
        updateCellData(tableroDocumentStore, rowIndex, columnIndex, newCellValue);
    });
};
const getCellValueAsObj = (val) => {
    return isValueComponentUrlType(val) ? { componentUrl: val } : { cellValue: val };
};
const isValueComponentUrlType = (val) => {
    // TODO discuss and change this as its not a perfect check for componentURl
    return Object(util__WEBPACK_IMPORTED_MODULE_1__["isString"])(val) && val.startsWith('/');
};
const updateCellData = (tableroDocumentStore, rowIndex, columnIndex, cellValue) => {
    if (cellValue === undefined) {
        return;
    }
    const currentCellValue = tableroDocumentStore.table.getCellValue(rowIndex, columnIndex);
    const newCellValue = JSON.stringify(getCellValueAsObj(cellValue));
    if (currentCellValue !== newCellValue) {
        tableroDocumentStore.table.setCellValue(rowIndex, columnIndex, newCellValue);
    }
};
const deleteTableRow = (tableroDocumentStore, data) => {
    const indexOfRow = tableroDocumentStore.rowIdToIndex[data.rowId];
    if (indexOfRow !== undefined) {
        Object(_TableroAdapter__WEBPACK_IMPORTED_MODULE_0__["deleteRow"])(tableroDocumentStore, indexOfRow);
    }
};
const getColumnData = (tableroDocumentStore) => {
    let columnData;
    columnData = [];
    for (let colIndex in tableroDocumentStore.colIdToIndex) {
        if (colIndex !== undefined)
            columnData.push(colIndex);
    }
    return columnData;
};
const getIndexFromId = (mapping, id) => {
    if (mapping === undefined || mapping[id] === undefined) {
        throw new Error('We expect that the mapping is always defined and the id exists.');
    }
    return mapping[id];
};
const getFluidUserWithDefaults = (componentContext, clientId) => {
    /**
     * A given user can have multiple clients open at the same time. The audience contains information about the user for each client.
     * The userId helps us identify a given user across all instances of the app.
     * Here, we extract the userId of the current author.
     */
    const fluidClient = componentContext.getAudience().getMember(clientId);
    // Don't show presence when the user is undefined.
    if (!fluidClient) {
        return undefined;
    }
    const user = fluidClient.user;
    // If we aren't able to get a username (happens for local-test-server and routerlicious) just show the client id as the username.
    const userId = user ? user.id : clientId;
    const userName = user && user.name ? user.name : clientId;
    return {
        id: userId,
        name: userName,
        email: user && user.email
    };
};


/***/ }),

/***/ "../tablero/lib/dataModel/TableroStore.js":
/*!************************************************!*\
  !*** ../tablero/lib/dataModel/TableroStore.js ***!
  \************************************************/
/*! exports provided: EditMode, TableroStore */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "EditMode", function() { return EditMode; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "TableroStore", function() { return TableroStore; });
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../telemetry */ "../tablero/lib/telemetry/ErrorEvents.js");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../telemetry */ "../tablero/lib/telemetry/utility.js");
/* harmony import */ var _fluidx_utilities__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @fluidx/utilities */ "../utilities/lib/logger/LoggerHelpers.js");


var EditMode;
(function (EditMode) {
    EditMode[EditMode["Undo"] = 0] = "Undo";
    EditMode[EditMode["Redo"] = 1] = "Redo";
    EditMode[EditMode["Edit"] = 2] = "Edit";
})(EditMode || (EditMode = {}));
// TODO:UNDO: revertibleArray should move into the transaction object since they are per transaction information.
//            (real transaction is not yet implemented, so we do not have a transaction object yet).
// TODO:UNDO: Should logger belong to TableroStore?
class TableroStore {
    constructor(tableroViewStore, tableroDocumentStore, logger) {
        this.tableroViewStore = tableroViewStore;
        this.tableroDocumentStore = tableroDocumentStore;
        this.logger = logger;
        this.revertInProgress = false;
        this.undoHandler = undefined;
    }
    revert(revertibleItem, isUndo) {
        this.postTransaction(() => {
            try {
                this.revertInProgress = true;
                const itemsToRevert = [...revertibleItem];
                while (itemsToRevert.length > 0) {
                    const item = itemsToRevert.pop();
                    item.revert();
                }
            }
            catch (error) {
                // TODO:TABLERO: Log the error.
            }
            finally {
                this.revertInProgress = false;
            }
        }, revertibleItem.label, isUndo ? EditMode.Undo : EditMode.Redo, true /* setNoop */);
    }
    postTransaction(action, actionLabel, editMode, setNoop, requestOrigin) {
        // We are expecting nested transaction -- at least for now, so the cache should be clean at this moment.
        if (this.revertibleArray && this.revertibleArray.length > 0) {
            this.logger &&
                Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_2__["sendErrorEvent"])(this.logger, {
                    eventName: _telemetry__WEBPACK_IMPORTED_MODULE_0__["ErrorEvent"].TransactionCorruption,
                    errorCode: _telemetry__WEBPACK_IMPORTED_MODULE_0__["ErrorEvent"].TransactionCorruption
                });
            delete this.revertibleArray;
        }
        try {
            // Execute
            const maybePromise = action();
            if (maybePromise instanceof Promise) {
                maybePromise
                    .then(() => {
                    this.onTransactionCompleted(editMode, actionLabel, setNoop, requestOrigin);
                })
                    .catch((error) => {
                    this.logger && Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_2__["sendErrorEvent"])(this.logger, { eventName: _telemetry__WEBPACK_IMPORTED_MODULE_0__["ErrorEvent"].TransactionExecutionError }, error);
                    this.clearCache();
                });
            }
            else {
                this.onTransactionCompleted(editMode, actionLabel, setNoop, requestOrigin);
            }
        }
        catch (error) {
            this.logger && Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_2__["sendErrorEvent"])(this.logger, { eventName: _telemetry__WEBPACK_IMPORTED_MODULE_0__["ErrorEvent"].TransactionExecutionError }, error);
            this.clearCache();
        }
    }
    onTransactionCompleted(editMode, transactionLabel, setNoop, requestOrigin) {
        transactionLabel &&
            this.logger &&
            Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_2__["sendUserActionEvent"])(this.logger, Object.assign({ eventName: transactionLabel }, Object(_telemetry__WEBPACK_IMPORTED_MODULE_1__["getUserActionConfig"])(requestOrigin)));
        if (this.revertibleArray && this.revertibleArray.length > 0) {
            const revertibleArray = this.revertibleArray;
            revertibleArray.label = this.getRevertibleItemLabel(transactionLabel || '', editMode);
            const undoUnit = {
                revertibleItem: revertibleArray,
                handler: this
            };
            if (this.undoHandler === undefined) {
                return;
            }
            if (editMode === EditMode.Undo) {
                this.undoHandler.pushToRedoStack(undoUnit);
            }
            else if (editMode === EditMode.Redo) {
                this.undoHandler.pushToUndoStack(undoUnit);
            }
            else {
                this.undoHandler.pushToUndoStack(undoUnit);
                // Clear redo stack.
                while (!this.undoHandler.isRedoStackEmpty) {
                    const redoUnit = this.undoHandler.popRedoStack();
                    redoUnit.revertibleItem.dispose && redoUnit.revertibleItem.dispose();
                }
            }
            this.clearCache();
        }
        // Seems current Tablero code relies on this "noop" trick to trigger rendering, so this is the same trick to trigger
        // rendering for undo/redo.
        //
        // The reason we need this is during normal operation, we are changing the table structure before the noop, but during undo, we
        // are doing undo in the reverse order, so we the first op get undone is noop - which triggers rendering before the actual undo job.
        //
        // This line has to be here to make sure this op won't be part of the undo/redo stack.
        //
        // TODO:TABLERO: With this line we should be able to remove those 'noop' change in each individual proposal function.
        setNoop && this.tableroDocumentStore.rootMap.set('noopkey', 'noop');
    }
    clearCache() {
        delete this.revertibleArray;
    }
    getRevertibleItemLabel(transactionLabel, editMode) {
        switch (editMode) {
            case EditMode.Edit:
                return `Undo${transactionLabel}`;
            case EditMode.Redo:
                return `Undo${transactionLabel.startsWith('Redo') ? transactionLabel.substr(4) : transactionLabel}`;
            case EditMode.Undo:
                return `Redo${transactionLabel.startsWith('Undo') ? transactionLabel.substr(4) : transactionLabel}`;
            default:
                return transactionLabel;
        }
    }
}


/***/ }),

/***/ "../tablero/lib/dataModel/adapterComponents/TableroAdapterBooleanComponent.js":
/*!************************************************************************************!*\
  !*** ../tablero/lib/dataModel/adapterComponents/TableroAdapterBooleanComponent.js ***!
  \************************************************************************************/
/*! exports provided: TableroAdapterBooleanComponent */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "TableroAdapterBooleanComponent", function() { return TableroAdapterBooleanComponent; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _view_cells_CheckboxCell__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../../view/cells/CheckboxCell */ "../tablero/lib/view/cells/CheckboxCell.js");
/* harmony import */ var _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../../view/table/Table.props */ "../tablero/lib/view/table/Table.props.js");



class TableroAdapterBooleanComponent {
    constructor(value) {
        this.value = value;
    }
    get ComponentProduceString() {
        return this;
    }
    get ComponentProduceBoolean() {
        return this;
    }
    get IComponentAdapterVisual() {
        return this;
    }
    async getString() {
        return this.value.toString();
    }
    getBoolean() {
        return this.value;
    }
    getAdapterComponentVisual(props) {
        return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_view_cells_CheckboxCell__WEBPACK_IMPORTED_MODULE_1__["CheckboxCell"], { rowId: props.rowId, columnId: props.columnId, columnType: _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__["DataType"].Checkbox, value: this.value, onCellValueProposal: props.onCellValueProposal, hasLocalSelection: props.hasLocalSelection, onRowCompletionProposal: props.onRowCompletionProposal, isReadOnlyMode: props.isReadOnlyMode }));
    }
}


/***/ }),

/***/ "../tablero/lib/dataModel/adapterComponents/TableroAdapterComponentFactory.js":
/*!************************************************************************************!*\
  !*** ../tablero/lib/dataModel/adapterComponents/TableroAdapterComponentFactory.js ***!
  \************************************************************************************/
/*! exports provided: createAdapterComponent */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "createAdapterComponent", function() { return createAdapterComponent; });
/* harmony import */ var _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../../view/table/Table.props */ "../tablero/lib/view/table/Table.props.js");
/* harmony import */ var _TableroAdapterStringComponent__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./TableroAdapterStringComponent */ "../tablero/lib/dataModel/adapterComponents/TableroAdapterStringComponent.js");
/* harmony import */ var _TableroAdapterBooleanComponent__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./TableroAdapterBooleanComponent */ "../tablero/lib/dataModel/adapterComponents/TableroAdapterBooleanComponent.js");
/* harmony import */ var _TableroAdapterDateComponent__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./TableroAdapterDateComponent */ "../tablero/lib/dataModel/adapterComponents/TableroAdapterDateComponent.js");
/* harmony import */ var _TableroAdapterRichTextComponent__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./TableroAdapterRichTextComponent */ "../tablero/lib/dataModel/adapterComponents/TableroAdapterRichTextComponent.js");





/**
 * Creates a new tablero adapter component
 * @param columnType type of tablero adapter component
 * @param value value of tablero adapter component. If not provided, use a default value
 */
const createAdapterComponent = (columnType, value) => {
    switch (columnType) {
        case _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].Text:
            return new _TableroAdapterStringComponent__WEBPACK_IMPORTED_MODULE_1__["TableroAdapterStringComponent"](value && typeof value === 'string' ? value : '');
        case _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].Checkbox:
            return new _TableroAdapterBooleanComponent__WEBPACK_IMPORTED_MODULE_2__["TableroAdapterBooleanComponent"](value && typeof value === 'boolean' ? value : false);
        case _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].Date:
            return new _TableroAdapterDateComponent__WEBPACK_IMPORTED_MODULE_3__["TableroAdapterDateComponent"](value && typeof value === 'object' && value instanceof Date ? value : undefined);
        case _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].RichText:
        case _view_table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].SimpleRichText:
            return new _TableroAdapterRichTextComponent__WEBPACK_IMPORTED_MODULE_4__["TableroAdapterRichTextComponent"](value && typeof value === 'string' ? value : '');
    }
};


/***/ }),

/***/ "../tablero/lib/dataModel/adapterComponents/TableroAdapterDateComponent.js":
/*!*********************************************************************************!*\
  !*** ../tablero/lib/dataModel/adapterComponents/TableroAdapterDateComponent.js ***!
  \*********************************************************************************/
/*! exports provided: TableroAdapterDateComponent */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "TableroAdapterDateComponent", function() { return TableroAdapterDateComponent; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _view_cells_DateCell__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../../view/cells/DateCell */ "../tablero/lib/view/cells/DateCell.js");
/* harmony import */ var _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../../view/table/Table.props */ "../tablero/lib/view/table/Table.props.js");



class TableroAdapterDateComponent {
    constructor(value) {
        this.value = value;
    }
    get ComponentProduceString() {
        return this;
    }
    get ComponentProduceDate() {
        return this;
    }
    get IComponentAdapterVisual() {
        return this;
    }
    async getString() {
        return this.value ? this.value.toString() : '';
    }
    getDate() {
        return this.value;
    }
    getAdapterComponentVisual(props) {
        return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_view_cells_DateCell__WEBPACK_IMPORTED_MODULE_1__["DateCell"], { rowId: props.rowId, columnId: props.columnId, columnType: _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__["DataType"].Date, hasLocalSelection: props.hasLocalSelection, onCellValueProposal: props.onCellValueProposal, cellWrapperFocused: props.cellWrapperFocused, notifyCellWrapperFocusExecuted: props.notifyCellWrapperFocusExecuted, value: this.value || '', componentContext: props.componentContext, isReadOnlyMode: props.isReadOnlyMode }));
    }
}


/***/ }),

/***/ "../tablero/lib/dataModel/adapterComponents/TableroAdapterRichTextComponent.js":
/*!*************************************************************************************!*\
  !*** ../tablero/lib/dataModel/adapterComponents/TableroAdapterRichTextComponent.js ***!
  \*************************************************************************************/
/*! exports provided: TableroAdapterRichTextComponent */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "TableroAdapterRichTextComponent", function() { return TableroAdapterRichTextComponent; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _view_cells_RichTextCell__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../../view/cells/RichTextCell */ "../tablero/lib/view/cells/RichTextCell.js");


class TableroAdapterRichTextComponent {
    constructor(value) {
        this.value = value;
    }
    get ComponentRichTextEditor() {
        return this;
    }
    get IComponentAdapterVisual() {
        return this;
    }
    getText(_startOffset, _endOffset) {
        return this.value;
    }
    getHtml(_startOffset, _endOffset) {
        return this.value;
    }
    setText(text) {
        this.value = text;
    }
    setHtml(html) {
        this.value = html;
    }
    getAdapterComponentVisual(props) {
        return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_view_cells_RichTextCell__WEBPACK_IMPORTED_MODULE_1__["RichTextCell"], { rowId: props.rowId, columnId: props.columnId, columnType: props.columnType, hasLocalSelection: props.hasLocalSelection, onCellValueProposal: props.onCellValueProposal, value: this.value || '', componentContext: props.componentContext, isReadOnlyMode: props.isReadOnlyMode }));
    }
}


/***/ }),

/***/ "../tablero/lib/dataModel/adapterComponents/TableroAdapterStringComponent.js":
/*!***********************************************************************************!*\
  !*** ../tablero/lib/dataModel/adapterComponents/TableroAdapterStringComponent.js ***!
  \***********************************************************************************/
/*! exports provided: TableroAdapterStringComponent */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "TableroAdapterStringComponent", function() { return TableroAdapterStringComponent; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _view_cells_TextCell__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../../view/cells/TextCell */ "../tablero/lib/view/cells/TextCell.js");
/* harmony import */ var _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../../view/table/Table.props */ "../tablero/lib/view/table/Table.props.js");



class TableroAdapterStringComponent {
    constructor(value) {
        this.value = value;
    }
    get ComponentProduceString() {
        return this;
    }
    get IComponentAdapterVisual() {
        return this;
    }
    async getString() {
        return this.value;
    }
    getAdapterComponentVisual(props) {
        return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_view_cells_TextCell__WEBPACK_IMPORTED_MODULE_1__["TextCell"], { rowId: props.rowId, columnId: props.columnId, columnType: _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__["DataType"].Text, value: this.value, onCellValueProposal: props.onCellValueProposal, hasLocalSelection: props.hasLocalSelection, columnWidth: props.width, onUpdateCellSelection: props.onUpdateCellSelection, sendInsertComponentRequestData: props.sendInsertComponentRequestData, enableUndoRedo: props.enableUndoRedo, componentContext: props.componentContext, logger: props.logger, onMoveSelectionFromCell: props.onMoveSelectionFromCell, moveSelectionToCellDetails: props.moveSelectionToCellDetails, isReadOnlyMode: props.isReadOnlyMode }));
    }
}


/***/ }),

/***/ "../tablero/lib/intelligence/InsertComponentRequest.js":
/*!*************************************************************!*\
  !*** ../tablero/lib/intelligence/InsertComponentRequest.js ***!
  \*************************************************************/
/*! exports provided: notifyInsertComponentRequest */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "notifyInsertComponentRequest", function() { return notifyInsertComponentRequest; });
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../telemetry */ "../tablero/lib/telemetry/ErrorEvents.js");
/* harmony import */ var _fluidx_utilities__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @fluidx/utilities */ "../utilities/lib/logger/LoggerHelpers.js");
/* harmony import */ var _utilities_createComponent__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../utilities/createComponent */ "../tablero/lib/utilities/createComponent.js");



/**
 * Notifies consumers that the host requests to insert a component.
 * @param context The context that holds all the properties needed for notifying consumers about the opportunity to insert a component.
 */
const notifyInsertComponentRequest = (context) => {
    const onInsert = (containerConfiguration) => {
        if (containerConfiguration.configurationType === 'componentConfiguration' &&
            containerConfiguration.loadComponentData.componentType) {
            // tslint:disable-next-line: no-floating-promises
            Object(_utilities_createComponent__WEBPACK_IMPORTED_MODULE_2__["createComponentWithRegistryType"])(containerConfiguration.loadComponentData.componentType, context.componentContext.containerRuntime).then((result) => {
                if (result) {
                    const cellContext = context.cellContext;
                    cellContext.onCellValueProposal(cellContext.rowId, cellContext.columnId, result);
                }
                else {
                    context.logger &&
                        Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_1__["sendErrorEvent"])(context.logger, {
                            eventName: _telemetry__WEBPACK_IMPORTED_MODULE_0__["ErrorEvent"].ComponentNestingError,
                            errorCode: _telemetry__WEBPACK_IMPORTED_MODULE_0__["ErrorCode"].ComponentNestingError.FailedToNest
                        });
                }
            });
        }
    };
    const onClose = (replaceWithExistingContent, existingContent, setFocus) => {
        if (context) {
            if (replaceWithExistingContent && context.cellContext.onReplaceWhenClosed) {
                context.cellContext.onReplaceWhenClosed(existingContent);
            }
            const hostingRef = context.cellContext.hostingRef;
            const moveFocusToHost = context.cellContext.onMoveFocusToHostingCanvas;
            if (setFocus && hostingRef && hostingRef.current) {
                hostingRef.current.focus();
                if (moveFocusToHost) {
                    moveFocusToHost();
                }
            }
        }
    };
    const userInput = context.hostContext.userInput;
    // We only want to send the value out to consumers if there's no input or is a single character in the cell.
    if (userInput && userInput.length !== 1) {
        return;
    }
    const insertComponentRequestData = {
        insertionContext: { onInsert, onClose },
        hostContext: context.hostContext
    };
    context.sendInsertComponentRequestData(insertComponentRequestData);
};


/***/ }),

/***/ "../tablero/lib/telemetry/Activities.js":
/*!**********************************************!*\
  !*** ../tablero/lib/telemetry/Activities.js ***!
  \**********************************************/
/*! exports provided: Activity */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Activity", function() { return Activity; });
/**
 * Long running tasks in Tablero. These have the potential for success / failure
 */
var Activity;
(function (Activity) {
    Activity["GetTableBodyData"] = "GetTableBodyData";
    Activity["RenderTableroView"] = "RenderTableroView";
})(Activity || (Activity = {}));


/***/ }),

/***/ "../tablero/lib/telemetry/ErrorEvents.js":
/*!***********************************************!*\
  !*** ../tablero/lib/telemetry/ErrorEvents.js ***!
  \***********************************************/
/*! exports provided: ErrorEvent, ErrorCode */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ErrorEvent", function() { return ErrorEvent; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ErrorCode", function() { return ErrorCode; });
/**
 * Define errors here and use these for logging.
 */
var ErrorEvent;
(function (ErrorEvent) {
    ErrorEvent["Assertion"] = "Assertion";
    ErrorEvent["Exception"] = "Exception";
    ErrorEvent["BackCompatibility"] = "BackCompatibility";
    ErrorEvent["ComponentNestingError"] = "ComponentNestingError";
    ErrorEvent["DataInitializationError"] = "DataInitializationError";
    ErrorEvent["CommandingSurfaceError"] = "CommandingSurfaceError";
    // Delete Row Column
    ErrorEvent["DeletedRowMissingError"] = "DeletedRowMissingError";
    ErrorEvent["DeletedColumnMissingError"] = "DeletedColumnMissingError";
    // Undo/Redo
    ErrorEvent["TransactionCorruption"] = "TransactionCorruption";
    ErrorEvent["TransactionExecutionError"] = "TransactionExecutionError";
    // Presence
    ErrorEvent["PresenceSignalSubmissionError"] = "PresenceSignalSubmissionError";
    ErrorEvent["PresenceManagerRequestError"] = "PresenceManagerRequestError";
    // UndoStack
    ErrorEvent["UndoStackLoadError"] = "UndoStackLoadError";
    // Default Component
    // This is used for retrieving the Last Edited Component off of the default component.
    ErrorEvent["DefaultComponentForLastEditedRequestError"] = "DefaultComponentForLastEditedRequestError";
})(ErrorEvent || (ErrorEvent = {}));
const ErrorCode = {
    DataInitializationError: {
        ErrorParsingCellValue: 'ErrorParsingCellValue'
    },
    ComponentNestingError: {
        FailedToNest: 'FailedToNest',
        NestFailedToLoad: 'NestFailedToLoad',
        NestHasNoRenderableVisuals: 'NestHasNoRenderableVisuals',
        NestIsNotAValidProducer: 'NestIsNotAValidProducer',
        AdapterComponentUndefined: 'AdapterComponentUndefined',
        CellDataUndefined: 'CellDataUndefined'
    },
    BackCompatibility: {
        SupportRootmapPresets: 'SupportRootmapPresets',
        FailedToGetRootmapPresets: 'FailedToGetRootmapPresets'
    }
};


/***/ }),

/***/ "../tablero/lib/telemetry/UserAction.js":
/*!**********************************************!*\
  !*** ../tablero/lib/telemetry/UserAction.js ***!
  \**********************************************/
/*! exports provided: UserAction */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "UserAction", function() { return UserAction; });
var UserAction;
(function (UserAction) {
    UserAction["CommitTableCellValue"] = "CommitTableCellValue";
    UserAction["InsertTableRow"] = "InsertTableRow";
    UserAction["InsertTableColumn"] = "InsertTableColumn";
    UserAction["DeleteTableRow"] = "DeleteTableRow";
    UserAction["DeleteTableColumn"] = "DeleteTableColumn";
    UserAction["ReorderTableColumns"] = "ReorderTableColumns";
    UserAction["ChangeTableColumnType"] = "ChangeTableColumnType";
    UserAction["ChangeTableColumnTitle"] = "ChangeTableColumnTitle";
    UserAction["UnhideTableColumns"] = "UnhideTableColumns";
    UserAction["SortTableByColumn"] = "SortTableByColumn";
    UserAction["FilterTableByColumn"] = "FilterTableByColumn";
    UserAction["ResizeTableColumn"] = "ResizeTableColumn";
    UserAction["ChangeTableTitle"] = "ChangeTableTitle";
    UserAction["FilterTableRows"] = "FilterTableRows";
    UserAction["ResizeTableColumnFromCommandingMenu"] = "ResizeTableColumnFromCommandingMenu";
})(UserAction || (UserAction = {}));


/***/ }),

/***/ "../tablero/lib/telemetry/UserActionRequestOrigin.js":
/*!***********************************************************!*\
  !*** ../tablero/lib/telemetry/UserActionRequestOrigin.js ***!
  \***********************************************************/
/*! exports provided: RequestOrigin */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "RequestOrigin", function() { return RequestOrigin; });
// Defines the possible origin of interactions for logging Telemetry events
var RequestOrigin;
(function (RequestOrigin) {
    RequestOrigin["Grabber"] = "Grabber";
    RequestOrigin["Default"] = "Default";
})(RequestOrigin || (RequestOrigin = {}));


/***/ }),

/***/ "../tablero/lib/telemetry/utility.js":
/*!*******************************************!*\
  !*** ../tablero/lib/telemetry/utility.js ***!
  \*******************************************/
/*! exports provided: getUserActionConfig, getConfigurationTelemetryName */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getUserActionConfig", function() { return getUserActionConfig; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getConfigurationTelemetryName", function() { return getConfigurationTelemetryName; });
/* harmony import */ var _UserActionRequestOrigin__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./UserActionRequestOrigin */ "../tablero/lib/telemetry/UserActionRequestOrigin.js");
/* harmony import */ var _configuration_Configuration__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../configuration/Configuration */ "../tablero/lib/configuration/Configuration.js");


const TelemetryComponentName = {
    [_configuration_Configuration__WEBPACK_IMPORTED_MODULE_1__["Configuration"].actionItemManager]: 'actionItemManager',
    [_configuration_Configuration__WEBPACK_IMPORTED_MODULE_1__["Configuration"].richTextTablero]: 'richTextTablero',
    [_configuration_Configuration__WEBPACK_IMPORTED_MODULE_1__["Configuration"].tasks]: 'tasks',
    [_configuration_Configuration__WEBPACK_IMPORTED_MODULE_1__["Configuration"].default]: 'plainTextTablero '
};
/**
 * returns extra info to be passed along with user action event.
 * Currently, Other than grabber, By default no extra info is passed for Tablero interactions.
 * TODO: Add other specifics for events (like: row insertion/deletion, column insertion/deletion, resizing etc.) triggered from commanding menu, Tab, mouse, keyboard actions.
 * It can be utilized to send the request origin and log activities from Planner/Task for futuristic purpose.
 * @param requestOrigin
 */
const getUserActionConfig = (requestOrigin) => {
    switch (requestOrigin) {
        case _UserActionRequestOrigin__WEBPACK_IMPORTED_MODULE_0__["RequestOrigin"].Grabber:
            return { triggerMethod: _UserActionRequestOrigin__WEBPACK_IMPORTED_MODULE_0__["RequestOrigin"].Grabber };
        default:
            return undefined;
    }
};
const getConfigurationTelemetryName = (configurationType) => {
    return TelemetryComponentName[configurationType] || '';
};


/***/ }),

/***/ "../tablero/lib/utilities/attributionUtils.js":
/*!****************************************************!*\
  !*** ../tablero/lib/utilities/attributionUtils.js ***!
  \****************************************************/
/*! exports provided: attributionEnabled, getAttributionDataWithAllyText */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "attributionEnabled", function() { return attributionEnabled; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getAttributionDataWithAllyText", function() { return getAttributionDataWithAllyText; });
/* harmony import */ var _fluidx_utilities__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @fluidx/utilities */ "../utilities/lib/settings/settingsValue.js");
/* harmony import */ var _dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../dataModel/TableroAdapter */ "../tablero/lib/dataModel/TableroAdapter.js");
/* harmony import */ var _view_table_TableAppFeatures__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../view/table/TableAppFeatures */ "../tablero/lib/view/table/TableAppFeatures.js");
/* harmony import */ var _view_StringTable__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../view/StringTable */ "../tablero/lib/view/StringTable.js");




const attributionEnabled = () => {
    return Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_0__["getSettingsValue"])({
        name: 'scriptor.attribution',
        defaultValue: Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_0__["enableUpToAudience"])('Production', _view_table_TableAppFeatures__WEBPACK_IMPORTED_MODULE_2__["audience"])
    });
};
/**
 * Retrieves attribution details from provided rows and column. If localPresence is in a cell,
 * that cell's data is retrieved and used to set accessibility text.
 */
const getAttributionDataWithAllyText = (tableroStore, rows, columns, hostUser, localPresence) => {
    const attributionMap = new Map();
    let textToSet = '';
    // Used to make sure only one user is marked for selection in the callout
    let userMarked = false;
    let userNameText = '';
    columns.map((column) => {
        rows.forEach((row) => {
            var _a;
            const shouldSetAccessibilityText = localPresence !== undefined &&
                localPresence.type === 'cell' &&
                localPresence.rowId === row.id &&
                localPresence.columnId === column.id;
            const componentAttributionData = (_a = row.data[column.id].ComponentProduceAttribution) === null || _a === void 0 ? void 0 : _a.getAttributionUXData();
            if (componentAttributionData !== undefined) {
                const componentUserData = componentAttributionData.attributionData;
                // For each user that attributed in a Scriptor cell, update Tablero's map with corresponding cell
                for (let index = 0; index < componentUserData.length; index += 1) {
                    const userInfo = componentUserData[index].userInfo;
                    // Add the current cell to user's attribution list
                    const attributionDetails = updateAttributionDataForCell(attributionMap.get(userInfo.id), userInfo, row.id, column.id, 
                    /*shouldSetSelection*/ !userMarked && shouldSetAccessibilityText);
                    // Make sure to mark only one user.
                    if (!userMarked && shouldSetAccessibilityText) {
                        userMarked = true;
                    }
                    // Update map
                    attributionMap.set(userInfo.id, attributionDetails);
                    // Add this user to accessibility text if local presence was in current cell.
                    if (shouldSetAccessibilityText) {
                        if (index === 0) {
                            // First user
                            userNameText = userInfo.name;
                        }
                        else if (index > 0 && index === componentUserData.length - 1) {
                            // Last user
                            userNameText += _view_StringTable__WEBPACK_IMPORTED_MODULE_3__["andSpace"] + userInfo.name;
                        }
                        else if (index > 0 && index < componentUserData.length - 1) {
                            // Users other than first and last separated by space
                            userNameText += _view_StringTable__WEBPACK_IMPORTED_MODULE_3__["commaSpace"] + userInfo.name;
                        }
                    }
                } // end component loop
                // Add the exact cell content to accessibility text after user names
                if (shouldSetAccessibilityText && componentAttributionData.getAccessibilityDivText) {
                    textToSet = _view_StringTable__WEBPACK_IMPORTED_MODULE_3__["attributionAllyText"]
                        .replace('{1}', userNameText)
                        .replace('{2}', componentAttributionData.getAccessibilityDivText());
                }
            }
            else {
                // TODO: Remove this backward compatibility for simple non-scriptor cells
                const fluidUser = Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_1__["getAttributionUserFromCellProperties"])(tableroStore.tableroDocumentStore, row.id, column.id);
                if (fluidUser) {
                    const attributionDetails = attributionMap.get(fluidUser.id);
                    const cellData = row.data[column.id];
                    if (shouldSetAccessibilityText && cellData.ComponentProduceString) {
                        cellData.ComponentProduceString.getString()
                            .then((value) => {
                            if (value) {
                                textToSet = _view_StringTable__WEBPACK_IMPORTED_MODULE_3__["attributionAllyText"].replace('{1}', fluidUser.name).replace('{2}', value);
                            }
                        })
                            .catch( /* do nothing*/);
                    }
                    attributionMap.set(fluidUser.id, updateAttributionDataForCell(attributionDetails, fluidUser, row.id, column.id, 
                    /*shouldSetSelection*/ !userMarked && shouldSetAccessibilityText));
                }
            }
        });
    });
    // If only self has attributed, skip showing attribution
    if (attributionMap.size === 1 && attributionMap.has(hostUser.id)) {
        return { attributionData: [], accessibilityText: '' };
    }
    const attributionUserList = [...attributionMap.values()];
    // If local presence was in table title to header row, announce the number of users that wrote in this table.
    if (textToSet.length === 0 && attributionUserList.length > 0) {
        textToSet = getAccessibilityDivTextForTableTitle(attributionUserList);
    }
    return { attributionData: attributionUserList, accessibilityText: textToSet };
};
function getAccessibilityDivTextForTableTitle(attributionData) {
    let userNameText = attributionData[0].userInfo.name;
    if (attributionData.length > 1) {
        userNameText = _view_StringTable__WEBPACK_IMPORTED_MODULE_3__["attributionUserListText"]
            .replace('{1}', attributionData[0].userInfo.name)
            .replace('{2}', `${attributionData.length - 1}`);
    }
    return _view_StringTable__WEBPACK_IMPORTED_MODULE_3__["attributionTableAllyText"].replace('{1}', userNameText);
}
function updateAttributionDataForCell(data, fluidUser, rowId, colId, shouldSetSelection) {
    let attributionDetails = data;
    if (attributionDetails === undefined) {
        attributionDetails = {
            userInfo: fluidUser,
            range: { type: 'cells', cells: {} }
        };
    }
    let cellRange = attributionDetails.range;
    if (cellRange && cellRange.type === 'cells') {
        let cols = cellRange.cells[rowId];
        if (cols === undefined) {
            cols = [];
        }
        cols.push(colId);
        cellRange.cells[rowId] = cols;
    }
    attributionDetails.range = cellRange;
    if (shouldSetSelection) {
        attributionDetails.shouldSetSelection = shouldSetSelection;
    }
    return attributionDetails;
}


/***/ }),

/***/ "../tablero/lib/utilities/createComponent.js":
/*!***************************************************!*\
  !*** ../tablero/lib/utilities/createComponent.js ***!
  \***************************************************/
/*! exports provided: createComponentWithRegistryType, createComponentWithDataProvider */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "createComponentWithRegistryType", function() { return createComponentWithRegistryType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "createComponentWithDataProvider", function() { return createComponentWithDataProvider; });
const createComponentWithRegistryType = (componentType, containerRuntime, initialConfig) => {
    let requestUrl = '_create?registryType=' + componentType; // TODO: make sure this works with string[] (spfx)
    return createComponent(requestUrl, containerRuntime, initialConfig);
};
const createComponentWithDataProvider = (dataProvider, containerRuntime, initialConfig) => {
    let requestUrl = '_create?dataProvider=' + dataProvider;
    return createComponent(requestUrl, containerRuntime, initialConfig);
};
const createComponent = async (requestUrl, containerRuntime, initialConfig) => {
    const response = await containerRuntime.request({ url: requestUrl, headers: { initialConfig } });
    if (response.status === 200 && response.mimeType === 'fluid/component' && response.value !== undefined) {
        return response.value;
    }
    return undefined;
};


/***/ }),

/***/ "../tablero/lib/utilities/stylingUtils.js":
/*!************************************************!*\
  !*** ../tablero/lib/utilities/stylingUtils.js ***!
  \************************************************/
/*! exports provided: listenOnThemeChange, clearAndUpdateThemeVariables, removeThemeEventListener */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "listenOnThemeChange", function() { return listenOnThemeChange; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "clearAndUpdateThemeVariables", function() { return clearAndUpdateThemeVariables; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "removeThemeEventListener", function() { return removeThemeEventListener; });
/*
  stylingUtils provides the generic methods that allows the components to attach 'themePropertiesChanged' event listener and
  resetThemeVariables on hosting elements of components as and when required.
  Later on listeners for font changes and other generic styling methods to be included here only.
*/
/**
 * It attaches the listener to respond whenever theme changes dynamically and resets the theme css properties.
 * @param stylingService The Styling container service that exposes application theming properties
 * @param listener callback function to be attached on themePropertiesChanged event
 */
const listenOnThemeChange = (stylingService, listener) => {
    stylingService && stylingService.on('themePropertiesChanged', listener);
};
/**
 * It removes the previously attached theme properties from root element and updates the new properties derieved from appStyleConfig.
 * @param previousThemeProperties - previously attached theme properties on the root element.
 * @param hostingElement root element from which theme properties will be removed and re-set.
 * @param stylingService The Styling container service that exposes application theming properties
 */
const clearAndUpdateThemeVariables = (previousThemeProperties, hostingElement, stylingService) => {
    if (!stylingService || !hostingElement) {
        return;
    }
    if (previousThemeProperties) {
        for (const previousThemeProperty of previousThemeProperties) {
            hostingElement.style.removeProperty(previousThemeProperty[0]);
        }
    }
    const themeProperties = stylingService.themeProperties();
    for (const themeProperty of themeProperties) {
        hostingElement.style.setProperty(themeProperty.name, themeProperty.value);
    }
};
/**
 * Removes the listener from the themePropertiesChanged event type.
 * @param stylingService The Styling container service that exposes application theming properties
 * @param listener callback function to be removed from themePropertiesChanged event
 */
const removeThemeEventListener = (stylingService, listener) => {
    stylingService && stylingService.removeListener('themePropertiesChanged', listener);
};


/***/ }),

/***/ "../tablero/lib/view/Constants.js":
/*!****************************************!*\
  !*** ../tablero/lib/view/Constants.js ***!
  \****************************************/
/*! exports provided: headerRowIndex, columnWidthIncrement */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "headerRowIndex", function() { return headerRowIndex; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "columnWidthIncrement", function() { return columnWidthIncrement; });
// Used to index the coauthor locations in header row
const headerRowIndex = 'HEADER_ROW_INDEX';
// Column width increment when resizing a column via keyboard.
const columnWidthIncrement = 10;


/***/ }),

/***/ "../tablero/lib/view/StringTable.js":
/*!******************************************!*\
  !*** ../tablero/lib/view/StringTable.js ***!
  \******************************************/
/*! exports provided: columnType, textType, checkBoxType, sortAscending, sortDescending, hideColumn, showAllColumns, filterRows, columnMenuSortLabel, columnMenuFilterLabel, checkedDisplayText, uncheckedDisplayText, noFilterMessage, sortDownLabel, sortUpLabel, headerCellAriaLabel, resizeColumnLabel, addTaskLabel, insertRowBelowLabel, insertTaskBelowLabel, insertRowAboveLabel, insertTaskAboveLabel, deleteRowLabel, deleteTaskLabel, resizeAnnoucement, resizeAnnoucementIncrease, resizeAnnoucementDecrease, attributionAllyText, attributionTableAllyText, attributionUserListText, commaSpace, andSpace, resizeTextStart, resizeTextLeftArrow, resizeTextMiddle, resizeTextRightArrow, resizeTextEnd */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "columnType", function() { return columnType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "textType", function() { return textType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "checkBoxType", function() { return checkBoxType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "sortAscending", function() { return sortAscending; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "sortDescending", function() { return sortDescending; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "hideColumn", function() { return hideColumn; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "showAllColumns", function() { return showAllColumns; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "filterRows", function() { return filterRows; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "columnMenuSortLabel", function() { return columnMenuSortLabel; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "columnMenuFilterLabel", function() { return columnMenuFilterLabel; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "checkedDisplayText", function() { return checkedDisplayText; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "uncheckedDisplayText", function() { return uncheckedDisplayText; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "noFilterMessage", function() { return noFilterMessage; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "sortDownLabel", function() { return sortDownLabel; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "sortUpLabel", function() { return sortUpLabel; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "headerCellAriaLabel", function() { return headerCellAriaLabel; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "resizeColumnLabel", function() { return resizeColumnLabel; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "addTaskLabel", function() { return addTaskLabel; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "insertRowBelowLabel", function() { return insertRowBelowLabel; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "insertTaskBelowLabel", function() { return insertTaskBelowLabel; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "insertRowAboveLabel", function() { return insertRowAboveLabel; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "insertTaskAboveLabel", function() { return insertTaskAboveLabel; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "deleteRowLabel", function() { return deleteRowLabel; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "deleteTaskLabel", function() { return deleteTaskLabel; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "resizeAnnoucement", function() { return resizeAnnoucement; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "resizeAnnoucementIncrease", function() { return resizeAnnoucementIncrease; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "resizeAnnoucementDecrease", function() { return resizeAnnoucementDecrease; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "attributionAllyText", function() { return attributionAllyText; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "attributionTableAllyText", function() { return attributionTableAllyText; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "attributionUserListText", function() { return attributionUserListText; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "commaSpace", function() { return commaSpace; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "andSpace", function() { return andSpace; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "resizeTextStart", function() { return resizeTextStart; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "resizeTextLeftArrow", function() { return resizeTextLeftArrow; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "resizeTextMiddle", function() { return resizeTextMiddle; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "resizeTextRightArrow", function() { return resizeTextRightArrow; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "resizeTextEnd", function() { return resizeTextEnd; });
// TODO: TABLERO: FIgure out how components will localize their content and load in the right strings.
const columnType = 'Type';
const textType = 'Text';
const checkBoxType = 'Checkbox';
const sortAscending = 'Ascending';
const sortDescending = 'Descending';
const hideColumn = 'Hide Column';
const showAllColumns = 'Show all Columns';
const filterRows = 'Filter';
const columnMenuSortLabel = 'Sort';
const columnMenuFilterLabel = 'Filter by Values';
const checkedDisplayText = 'Checked';
const uncheckedDisplayText = 'Unchecked';
const noFilterMessage = 'Filtering not supported';
const sortDownLabel = 'SortDown';
const sortUpLabel = 'SortUp';
const headerCellAriaLabel = 'Header Cell';
const resizeColumnLabel = 'Resize column';
const addTaskLabel = 'Add a task';
const insertRowBelowLabel = 'Insert row below';
const insertTaskBelowLabel = 'Insert task below';
const insertRowAboveLabel = 'Insert row above';
const insertTaskAboveLabel = 'Insert task above';
const deleteRowLabel = 'Delete row';
const deleteTaskLabel = 'Delete task';
const resizeAnnoucement = 'Width of column ';
const resizeAnnoucementIncrease = ' increased';
const resizeAnnoucementDecrease = ' decreased';
const attributionAllyText = '{1} wrote: {2}';
const attributionTableAllyText = '{1} contributed text in this table.';
const attributionUserListText = '{1} and {2} other';
const commaSpace = ', ';
const andSpace = ' and ';
const resizeTextStart = 'Use';
const resizeTextLeftArrow = 'left arrow key';
const resizeTextMiddle = 'or';
const resizeTextRightArrow = 'right arrow key';
const resizeTextEnd = 'to resize';


/***/ }),

/***/ "../tablero/lib/view/StylingConstants.js":
/*!***********************************************!*\
  !*** ../tablero/lib/view/StylingConstants.js ***!
  \***********************************************/
/*! exports provided: standardColors, standardCSSDisplayTypes, colorLabels, fontFamily, presenceBubbleHeight, presenceBubbleCollapsedWidth, attributedCellBackgroundForClick, attributedCellBackgroundForHover, cellWrapperBorderWidth, cellWrapperBorderWidthZoomLessThanOne, cellWrapperBorderWidthZoomLessThanHalf, cellWrapperBorderWidthZoomLessThanThird, cellWrapperBorderRadius, cellWrapperPaddingTopAndBottom, cellWrapperPaddingLeftAndRight, minComponentCellHeight, selectedCellBoxShadow, rowBottomBoxShadow, rowRightBoxShadow, rowLeftBoxShadow, resizeHandlerWidth, highContrastOutline, columnHCBoxShadow, rowHCBoxShadow, columnGrabberHCBoxShadow, grabberContainerHCOutline, titleFontSize, titleFontWeight, titleLineHeight, titleFontFamily, tableHeaderHeight, tableTitleMinWidth, tableScrollBarColor, tableScrollBarBorderRadius, tableScrollBarHeight, tableScrollBarWidth, rowGrabberWidth, columnGrabberHeight, columnGrabberMenuGap, rowGrabberMenuGap, columnGrabberMenuPadding, rowGrabberMenuPadding, grabberMenuItemHeight, grabberMenuItemWidth, grabberMenuItemFontSize, grabberMenuItemLineHeight, grabberBorderRadius, grabberMenuBorderRadius, addRowFontElementPadding, insertIconSize, deleteIconSize, sortIconSize, gripperIconSize, ellipseIconSize, circleIconSize, insertIconAnimationSpeed, tableBorderRadius, tableBottomMargin, tableRightMargin, tableRightPaddingGrabber, tableTextFontSize, tableTextLineHeight, tableTextFontFamily, tableTextFontWeight, defaultColumnWidth, minimumColumnWidth, tableroRightMargin, tableroLeftMargin, tableroBottomMargin, columnHeaderBottomElementHeight, columnHeaderFontSize, columnHeaderFontWeight, columnHeaderLineHeight, columnHeaderFontFamily, horizontalScrollBarGradientWidth, resizeColumnCalloutStyleConstants, dragPreviewOpacity */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "standardColors", function() { return standardColors; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "standardCSSDisplayTypes", function() { return standardCSSDisplayTypes; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "colorLabels", function() { return colorLabels; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "fontFamily", function() { return fontFamily; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "presenceBubbleHeight", function() { return presenceBubbleHeight; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "presenceBubbleCollapsedWidth", function() { return presenceBubbleCollapsedWidth; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "attributedCellBackgroundForClick", function() { return attributedCellBackgroundForClick; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "attributedCellBackgroundForHover", function() { return attributedCellBackgroundForHover; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "cellWrapperBorderWidth", function() { return cellWrapperBorderWidth; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "cellWrapperBorderWidthZoomLessThanOne", function() { return cellWrapperBorderWidthZoomLessThanOne; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "cellWrapperBorderWidthZoomLessThanHalf", function() { return cellWrapperBorderWidthZoomLessThanHalf; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "cellWrapperBorderWidthZoomLessThanThird", function() { return cellWrapperBorderWidthZoomLessThanThird; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "cellWrapperBorderRadius", function() { return cellWrapperBorderRadius; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "cellWrapperPaddingTopAndBottom", function() { return cellWrapperPaddingTopAndBottom; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "cellWrapperPaddingLeftAndRight", function() { return cellWrapperPaddingLeftAndRight; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "minComponentCellHeight", function() { return minComponentCellHeight; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "selectedCellBoxShadow", function() { return selectedCellBoxShadow; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "rowBottomBoxShadow", function() { return rowBottomBoxShadow; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "rowRightBoxShadow", function() { return rowRightBoxShadow; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "rowLeftBoxShadow", function() { return rowLeftBoxShadow; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "resizeHandlerWidth", function() { return resizeHandlerWidth; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "highContrastOutline", function() { return highContrastOutline; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "columnHCBoxShadow", function() { return columnHCBoxShadow; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "rowHCBoxShadow", function() { return rowHCBoxShadow; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "columnGrabberHCBoxShadow", function() { return columnGrabberHCBoxShadow; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "grabberContainerHCOutline", function() { return grabberContainerHCOutline; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "titleFontSize", function() { return titleFontSize; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "titleFontWeight", function() { return titleFontWeight; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "titleLineHeight", function() { return titleLineHeight; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "titleFontFamily", function() { return titleFontFamily; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableHeaderHeight", function() { return tableHeaderHeight; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableTitleMinWidth", function() { return tableTitleMinWidth; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableScrollBarColor", function() { return tableScrollBarColor; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableScrollBarBorderRadius", function() { return tableScrollBarBorderRadius; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableScrollBarHeight", function() { return tableScrollBarHeight; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableScrollBarWidth", function() { return tableScrollBarWidth; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "rowGrabberWidth", function() { return rowGrabberWidth; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "columnGrabberHeight", function() { return columnGrabberHeight; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "columnGrabberMenuGap", function() { return columnGrabberMenuGap; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "rowGrabberMenuGap", function() { return rowGrabberMenuGap; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "columnGrabberMenuPadding", function() { return columnGrabberMenuPadding; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "rowGrabberMenuPadding", function() { return rowGrabberMenuPadding; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "grabberMenuItemHeight", function() { return grabberMenuItemHeight; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "grabberMenuItemWidth", function() { return grabberMenuItemWidth; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "grabberMenuItemFontSize", function() { return grabberMenuItemFontSize; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "grabberMenuItemLineHeight", function() { return grabberMenuItemLineHeight; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "grabberBorderRadius", function() { return grabberBorderRadius; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "grabberMenuBorderRadius", function() { return grabberMenuBorderRadius; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "addRowFontElementPadding", function() { return addRowFontElementPadding; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "insertIconSize", function() { return insertIconSize; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "deleteIconSize", function() { return deleteIconSize; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "sortIconSize", function() { return sortIconSize; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "gripperIconSize", function() { return gripperIconSize; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ellipseIconSize", function() { return ellipseIconSize; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "circleIconSize", function() { return circleIconSize; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "insertIconAnimationSpeed", function() { return insertIconAnimationSpeed; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableBorderRadius", function() { return tableBorderRadius; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableBottomMargin", function() { return tableBottomMargin; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableRightMargin", function() { return tableRightMargin; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableRightPaddingGrabber", function() { return tableRightPaddingGrabber; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableTextFontSize", function() { return tableTextFontSize; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableTextLineHeight", function() { return tableTextLineHeight; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableTextFontFamily", function() { return tableTextFontFamily; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableTextFontWeight", function() { return tableTextFontWeight; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "defaultColumnWidth", function() { return defaultColumnWidth; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "minimumColumnWidth", function() { return minimumColumnWidth; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableroRightMargin", function() { return tableroRightMargin; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableroLeftMargin", function() { return tableroLeftMargin; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableroBottomMargin", function() { return tableroBottomMargin; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "columnHeaderBottomElementHeight", function() { return columnHeaderBottomElementHeight; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "columnHeaderFontSize", function() { return columnHeaderFontSize; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "columnHeaderFontWeight", function() { return columnHeaderFontWeight; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "columnHeaderLineHeight", function() { return columnHeaderLineHeight; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "columnHeaderFontFamily", function() { return columnHeaderFontFamily; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "horizontalScrollBarGradientWidth", function() { return horizontalScrollBarGradientWidth; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "resizeColumnCalloutStyleConstants", function() { return resizeColumnCalloutStyleConstants; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "dragPreviewOpacity", function() { return dragPreviewOpacity; });
/* harmony import */ var _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @fluidx/office-fluid-types */ "../office-fluid-types/lib/BohemiaScope/StylingUtilities/ThemeColorPalette.js");
/* harmony import */ var _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @fluidx/office-fluid-types */ "../office-fluid-types/lib/BohemiaScope/StylingUtilities/ThemingProperties.js");

const standardColors = {
    transparent: 'transparent',
    textColor: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_0__["ThemeColorPalette"].Neutral.neutralPrimary
};
const standardCSSDisplayTypes = {
    inline: 'inline',
    flex: 'flex',
    inlineBlock: 'inline-block'
};
const colorLabels = {
    tableBackgroundColor: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].VariantBorder}, #DDDDDD)`,
    tableWrapperBackgroundColor: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].BodyBackground}, ${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_0__["ThemeColorPalette"].Neutral.white})`,
    tableTitleText: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].BodyText}, ${standardColors.textColor})`,
    tableTitleBackground: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].BodyBackground}, ${standardColors.transparent})`,
    cellText: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].BodyText}, ${standardColors.textColor})`,
    cellBackground: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].BodyBackground}, ${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_0__["ThemeColorPalette"].Neutral.white})`,
    attributedCellBorderColor: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].FocusBorder}, #026367)`,
    headerRowBottomBorderColor: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].VariantBorderHovered}, #AAAAAA)`,
    localSelection: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].InputBackgroundChecked}, ${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_0__["ThemeColorPalette"].Accent.orange})`,
    localHover: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].InputPlaceholderBackgroundChecked}, #E8825D)`,
    presenceBubbleUserInitialsText: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].BodySubtext}, ${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_0__["ThemeColorPalette"].Neutral.white})`,
    checkbox: {
        background: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].InputBackground}, ${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_0__["ThemeColorPalette"].Neutral.white})`,
        colorUnchecked: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].InputBackgroundCheckedHovered}, ${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_0__["ThemeColorPaletteShared"].Gray180}`,
        colorUncheckedHovered: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].InputBackgroundChecked}, ${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_0__["ThemeColorPalette"].Accent.orange}`,
        colorChecked: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].InputBackgroundChecked}, ${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_0__["ThemeColorPalette"].Accent.orange})`,
        colorCheckedHovered: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].InputBackgroundChecked}, ${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_0__["ThemeColorPalette"].Accent.orange})`
    },
    grabber: {
        clickedBorder: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].InputBorder}, ${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_0__["ThemeColorPalette"].Accent.orange})`,
        clickedBackground: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].InputBorder}, ${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_0__["ThemeColorPalette"].Accent.orange})`,
        fontIcon: {
            color: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].ButtonText}, #AAAAAA)`,
            clickedColor: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].ButtonTextPressed}, ${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_0__["ThemeColorPalette"].Neutral.white})`,
            background: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].InputBorder}, ${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_0__["ThemeColorPalette"].Accent.orange})`
        }
    },
    menu: {
        background: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].MenuBackground}, ${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_0__["ThemeColorPalette"].Neutral.white})`,
        backgroundHover: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].MenuItemBackgroundHovered}, ${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_0__["ThemeColorPalette"].Neutral.neutralLighter})`,
        color: `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeColorPropertyNames"].MenuItemText}, ${standardColors.textColor})`
    }
};
const fontFamily = 'Avenir Next LT Pro, Avenir Next LT Pro_MSFontService, Segoe UI, Segoe UI_MSFontService, sans-serif';
// Presence container properties
const presenceBubbleHeight = 20;
const presenceBubbleCollapsedWidth = 20;
// Attribution highlight colors
const attributedCellBackgroundForClick = `rgb(194, 234, 238)`;
const attributedCellBackgroundForHover = `rgb(236, 251, 252)`;
// Table cell properties
const cellWrapperBorderWidth = 1;
const cellWrapperBorderWidthZoomLessThanOne = 2;
const cellWrapperBorderWidthZoomLessThanHalf = 3;
const cellWrapperBorderWidthZoomLessThanThird = 4;
const cellWrapperBorderRadius = 1;
const cellWrapperPaddingTopAndBottom = 10;
const cellWrapperPaddingLeftAndRight = 12;
const minComponentCellHeight = 20;
const selectedCellBoxShadow = '0px 0.6px 1.8px rgba(0, 0, 0, 0.108), 0px 3.2px 7.2px rgba(0, 0, 0, 0.132)';
const rowBottomBoxShadow = `0px 3.8px 5.8px -1px rgba(0, 0, 0, 0.132)`;
const rowRightBoxShadow = `3.8px 0px 5.8px -1px rgba(0,0,0,0.132)`;
const rowLeftBoxShadow = `-3.8px 0px 5.8px -1px rgba(0, 0, 0, 0.132)`;
const resizeHandlerWidth = 3;
// Table High Contrast properties
const highContrastOutline = '1px dashed window';
const columnHCBoxShadow = '0 -1px 0 0 highlight';
const rowHCBoxShadow = '1px 0 0 0 highlight';
const columnGrabberHCBoxShadow = '0 1px 0 0 highlight';
const grabberContainerHCOutline = '1px dashed highlight';
// Table title properties
const titleFontSize = `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeFontsCssPropertyName"].FontSizeHeading2}, 24px)`;
// TODO: Replace the font weight for `var(${ThemeFontsCssPropertyName.FontWeightHeading2}, 600)`.
// mergeStyles, where this is being used, complains about using something that's not a raw fontWeight value.
const titleFontWeight = 600;
const titleLineHeight = `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeLayoutCssPropertyName"].LineHeightHeading2}, 32px)`;
const titleFontFamily = `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeFontsCssPropertyName"].FontFamilyHeading2}, Avenir Next LT Pro Bold, Avenir Next LT Pro Bold_MSFontService, Avenir Next LT Pro, Avenir Next LT Pro_MSFontService, Segoe UI, Segoe UI_MSFontService, sans-serif)`;
// The height of the table header was set to 49 because it misbehaved with SPO paste in pages.
// Updated to 28 as per new design for Build 2020
const tableHeaderHeight = 28;
const tableTitleMinWidth = 60;
// Scroll bar properties
const tableScrollBarColor = _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_0__["ThemeColorPaletteShared"].Gray100;
const tableScrollBarBorderRadius = 3;
const tableScrollBarHeight = 6;
const tableScrollBarWidth = 6;
// Grabber properties
const rowGrabberWidth = 18;
const columnGrabberHeight = 18;
const columnGrabberMenuGap = 11;
const rowGrabberMenuGap = 11;
const columnGrabberMenuPadding = 4;
const rowGrabberMenuPadding = 4;
const grabberMenuItemHeight = 32;
const grabberMenuItemWidth = 32;
const grabberMenuItemFontSize = 16;
const grabberMenuItemLineHeight = 19;
const grabberBorderRadius = 2;
const grabberMenuBorderRadius = 2;
const addRowFontElementPadding = 1;
// Grabber icon properties (Fabric MDL2 icons)
const insertIconSize = 16; // F2E3 CircleAddition, F2E4 CircleAdditionSolid
const deleteIconSize = 20; // E74D Delete
const sortIconSize = 20; // EE68 SortUp, EE69 SortDown
const gripperIconSize = 18; // F772 GripperDotsVertical
const ellipseIconSize = 3; // F137 StatusCircleInner
const circleIconSize = 6; //  CircleRing
const insertIconAnimationSpeed = 100;
// Table properties
const tableBorderRadius = 2;
// Margin for scroll bar and bottom shadow
const tableBottomMargin = 7;
const tableRightMargin = 1;
const tableRightPaddingGrabber = 9;
const tableTextFontSize = `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeFontsCssPropertyName"].FontSizeBody}, 16px)`;
const tableTextLineHeight = `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeLayoutCssPropertyName"].LineHeightBody}, 20px)`;
const tableTextFontFamily = `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeFontsCssPropertyName"].FontFamilyBody}, Avenir Next LT Pro, Avenir Next LT Pro_MSFontService, Segoe UI, Segoe UI_MSFontService, sans-serif)`;
// TODO: Replace the font weight for `var(${ThemeFontsCssPropertyName.FontWeightBody}, 400)`.
// mergeStyles, where this is being used, complains about using something that's not a raw fontWeight value.
const tableTextFontWeight = 400;
const defaultColumnWidth = 170;
const minimumColumnWidth = 80;
const tableroRightMargin = 20;
const tableroLeftMargin = 20;
const tableroBottomMargin = 5;
// Column header properties
const columnHeaderBottomElementHeight = 1;
const columnHeaderFontSize = `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeFontsCssPropertyName"].FontSizeHeading3}, 16px)`;
// TODO: Replace the font weight for `var(${ThemeFontsCssPropertyName.FontWeightHeading3}, 600)`.
// mergeStyles, where this is being used, complains about using something that's not a raw fontWeight value.
const columnHeaderFontWeight = 600;
const columnHeaderLineHeight = `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeLayoutCssPropertyName"].LineHeightHeading3}, 22px)`;
const columnHeaderFontFamily = `var(${_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["ThemeFontsCssPropertyName"].FontFamilyHeading3}, Avenir Next LT Pro Demi, Avenir Next LT Pro Demi_MSFontService, Avenir Next LT Pro, Avenir Next LT Pro_MSFontService, Segoe UI, Segoe UI_MSFontService, sans-serif)`;
// Horizontal scroll bar gradient
const horizontalScrollBarGradientWidth = 12;
// Resize Column Callout style constants
const resizeColumnCalloutStyleConstants = {
    // Icon styles
    replyAltIconMarginLeftRight: 4,
    replyAltIconMarginTop: 3,
    replyAltIconMarginBottom: 1,
    replyAltIconVerticalAlign: 'top',
    replyAltIconSize: 12,
    replyAltIconTransform: 'rotate(180deg)',
    // Callout styles
    height: 28,
    beakWidth: 8.5,
    gapSpace: 8,
    backgroundColor: '#FFFFFF',
    boxShadow: '0px 1.6px 3.6px rgba(0, 0, 0, 0.13), 0px 0.3px 0.9px rgba(0, 0, 0, 0.11)',
    borderRadius: 4,
    // Text holder styles
    textHolderMargin: 0,
    textHolderPaddingLeftRight: 12,
    textHolderPaddingTopBottom: 6,
    // Text styles
    textMarginPadding: 0,
    textLineHeight: 16,
    textFontFamily: 'Segoe UI',
    textColor: '#252525'
};
// Drag preview.
const dragPreviewOpacity = 0.3;


/***/ }),

/***/ "../tablero/lib/view/cells/CellDefinitions.js":
/*!****************************************************!*\
  !*** ../tablero/lib/view/cells/CellDefinitions.js ***!
  \****************************************************/
/*! exports provided: CellStyleOverrideOption */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "CellStyleOverrideOption", function() { return CellStyleOverrideOption; });
var CellStyleOverrideOption;
(function (CellStyleOverrideOption) {
    CellStyleOverrideOption[CellStyleOverrideOption["Default"] = 0] = "Default";
    CellStyleOverrideOption[CellStyleOverrideOption["MaintainOldOverride"] = 1] = "MaintainOldOverride";
})(CellStyleOverrideOption || (CellStyleOverrideOption = {}));


/***/ }),

/***/ "../tablero/lib/view/cells/CellHelpers.js":
/*!************************************************!*\
  !*** ../tablero/lib/view/cells/CellHelpers.js ***!
  \************************************************/
/*! exports provided: getCellId, getCellStyleOverridesForRow, getCellStyleOverridesForCell, setAriaLabel */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getCellId", function() { return getCellId; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getCellStyleOverridesForRow", function() { return getCellStyleOverridesForRow; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getCellStyleOverridesForCell", function() { return getCellStyleOverridesForCell; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "setAriaLabel", function() { return setAriaLabel; });
/**
 * Return a unique string identifier for a table cell.
 * rowId is undefined for a column header cell.
 * @param columnId - column id of table cell.
 * @param rowId - row id of table cell.
 */
function getCellId(columnId, rowId) {
    return `${columnId}-${rowId ? rowId : 'header'}`;
}
function getCellStyleOverridesForRow(cellStyleOverride, rowId) {
    if (!cellStyleOverride) {
        return undefined;
    }
    let allEdges;
    // Don't filter out if rowId is not there (it's a column style and will get filtered at cell level)
    if (cellStyleOverride.allEdges && (!cellStyleOverride.allEdges.rowId || cellStyleOverride.allEdges.rowId === rowId)) {
        allEdges = cellStyleOverride.allEdges;
    }
    let singleEdge;
    if (cellStyleOverride.singleEdge &&
        (!cellStyleOverride.singleEdge.rowId || cellStyleOverride.singleEdge.rowId === rowId)) {
        singleEdge = cellStyleOverride.singleEdge;
    }
    let opacity;
    if (cellStyleOverride.opacity && (!cellStyleOverride.opacity.rowId || cellStyleOverride.opacity.rowId === rowId)) {
        opacity = cellStyleOverride.opacity;
    }
    if (allEdges || singleEdge || opacity) {
        return { allEdges, singleEdge, opacity };
    }
    return undefined;
}
function getCellStyleOverridesForCell(cellStyleOverride, columnId, rowId) {
    if (!cellStyleOverride) {
        return undefined;
    }
    let allEdges;
    // If row Id exists on override, it should be same as cell's row id
    // If column Id exists on override, it should be same as cell's column Id
    if (cellStyleOverride.allEdges &&
        (!cellStyleOverride.allEdges.rowId || cellStyleOverride.allEdges.rowId === rowId) &&
        (!cellStyleOverride.allEdges.columnId || cellStyleOverride.allEdges.columnId === columnId)) {
        allEdges = cellStyleOverride.allEdges;
    }
    let singleEdge;
    if (cellStyleOverride.singleEdge &&
        (!cellStyleOverride.singleEdge.rowId || cellStyleOverride.singleEdge.rowId === rowId) &&
        (!cellStyleOverride.singleEdge.columnId || cellStyleOverride.singleEdge.columnId === columnId)) {
        singleEdge = cellStyleOverride.singleEdge;
    }
    let opacity;
    if (cellStyleOverride.opacity &&
        (!cellStyleOverride.opacity.rowId || cellStyleOverride.opacity.rowId === rowId) &&
        (!cellStyleOverride.opacity.columnId || cellStyleOverride.opacity.columnId === columnId)) {
        opacity = cellStyleOverride.opacity;
    }
    if (allEdges || singleEdge || opacity) {
        return { allEdges, singleEdge, opacity };
    }
    return undefined;
}
const setAriaLabel = (target, label) => {
    if (target && label) {
        target.setAttribute('aria-label', label);
    }
};


/***/ }),

/***/ "../tablero/lib/view/cells/CellInput.js":
/*!**********************************************!*\
  !*** ../tablero/lib/view/cells/CellInput.js ***!
  \**********************************************/
/*! exports provided: CellInput */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "CellInput", function() { return CellInput; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/KeyCodes.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/styling/7.13.5_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/styling/lib/index.js");
/* harmony import */ var _StylingConstants__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../StylingConstants */ "../tablero/lib/view/StylingConstants.js");
/* harmony import */ var _intelligence_InsertComponentRequest__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../../intelligence/InsertComponentRequest */ "../tablero/lib/intelligence/InsertComponentRequest.js");
/* harmony import */ var _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! @fluidx/office-fluid-types */ "../office-fluid-types/lib/ComponentContracts/ComponentSelection.js");
/* harmony import */ var _utils_Helper__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../utils/Helper */ "../tablero/lib/view/utils/Helper.js");
/* harmony import */ var _CellHelpers__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ./CellHelpers */ "../tablero/lib/view/cells/CellHelpers.js");
/* harmony import */ var _table_TableAppFeatures__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../table/TableAppFeatures */ "../tablero/lib/view/table/TableAppFeatures.js");








// These are the events we don't want to allow to propagate since we want to handle them entirely inside the input
const keyCodesThatRemoveFocus = new Set([office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["KeyCodes"].enter, office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["KeyCodes"].escape, office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["KeyCodes"].tab]);
// Describes the various modes the text area can be in
var InputMode;
(function (InputMode) {
    // This is the state when the cell input is completely unfocused.
    InputMode[InputMode["unfocused"] = 0] = "unfocused";
    // This is when the cell is first selected and typing replaces the current cell input.
    InputMode[InputMode["invisibleOvertype"] = 1] = "invisibleOvertype";
    // This is when a second click is made on the cell to enter IP edit mode.
    InputMode[InputMode["ipEditMode"] = 2] = "ipEditMode";
})(InputMode || (InputMode = {}));
// Manages updates to the state of the text area
const getReducer = () => {
    return (state, action) => {
        switch (action.type) {
            case 'setInputValue': {
                const { inputValue } = action;
                return Object.assign(Object.assign({}, state), { inputValue });
            }
            case 'setInputMode': {
                const { inputMode } = action;
                return Object.assign(Object.assign({}, state), { inputMode });
            }
        }
    };
};
// Prevent content drag and drop for textarea. No such feature ask yet
const onDragAndDrop = (ev) => {
    ev.preventDefault();
};
/**
 * A React wrapper around HTML Input elements for use in the table cells. This has semantics of knowing when to commit
 * the value the user committed versus restoring the previous value.
 *
 * It is a controlled component, meaning we don't let the DOM control the state, we push state to the DOM.
 */
const CellInput = react__WEBPACK_IMPORTED_MODULE_0__["memo"](function CellInput(props) {
    const { rowId, columnId, columnType, onCellValueProposal, committedValue, onValueCommit, styleOverrides, hasLocalSelection, columnWidth, onUpdateCellSelection, readOnly, componentContext, enableUndoRedo, logger, sendInsertComponentRequestData, cellInputApi, onMoveSelectionFromCell, moveSelectionToCellDetails, additionalA11yLabel } = props;
    const textAreaRef = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](null);
    const divRef = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](null);
    if (cellInputApi) {
        cellInputApi.current = {
            focus: () => {
                var _a;
                (_a = textAreaRef.current) === null || _a === void 0 ? void 0 : _a.focus();
            }
        };
    }
    const [state, dispatch] = react__WEBPACK_IMPORTED_MODULE_0__["useReducer"](getReducer(), {
        inputValue: committedValue,
        inputMode: InputMode.unfocused
    });
    const { inputValue, inputMode } = state;
    const callOnMoveSelectionFromCell = (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        if (onMoveSelectionFromCell) {
            if (columnId !== undefined && rowId !== undefined) {
                // Since both rowId and columnId are present, it means the presence type is cell.
                onMoveSelectionFromCell({ type: 'cell', rowId, columnId })(ev);
            }
            else if (columnId !== undefined) {
                // Since only columnId is present, it means the presence type is column aka header cell.
                onMoveSelectionFromCell({ type: 'column', columnId })(ev);
            }
        }
    };
    // Handle cell caret enter in case of entry with moveSelectionToCellDetails.
    const onSelectionMovedToCell = (moveSelectionToCellDetails) => {
        if (textAreaRef.current) {
            textAreaRef.current.focus();
            switch (moveSelectionToCellDetails.direction) {
                case _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_5__["SelectionDirection"].left:
                case _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_5__["SelectionDirection"].up:
                case _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_5__["SelectionDirection"].down:
                    const len = textAreaRef.current.value.length;
                    textAreaRef.current.setSelectionRange(len, len);
                    break;
                case _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_5__["SelectionDirection"].right:
                    textAreaRef.current.setSelectionRange(0, 0);
                    break;
                default:
                    break;
            }
        }
    };
    const commitValueIfChanged = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((valueToCommit) => {
        // If the value to commit does not match the currently committed value, commit it!
        if (committedValue !== valueToCommit) {
            onValueCommit(valueToCommit);
        }
    }, [committedValue, onValueCommit]);
    const updateTextAreaHeight = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"](() => {
        // Update the height of the text area so as to support multi-line text in Tablero cells
        const textAreaRefHTMLElement = textAreaRef.current;
        if (textAreaRefHTMLElement) {
            // In order for us to also resize to smaller when the text shrinks set the height to auto first
            // so we don't continually grab a bigger scroll height
            textAreaRefHTMLElement.style.height = 'auto';
            textAreaRefHTMLElement.style.height =
                Math.max(textAreaRefHTMLElement.scrollHeight, _StylingConstants__WEBPACK_IMPORTED_MODULE_3__["minComponentCellHeight"]) + 'px';
        }
    }, [textAreaRef]);
    // TODO: Task 3720147: Bind to only allow row completion action on the AIM task cell
    const cellClassName = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_2__["mergeStyles"])({
        textDecoration: 'unset',
        width: '100%',
        border: 0,
        padding: '0px',
        boxSizing: 'border-box',
        fontFamily: _StylingConstants__WEBPACK_IMPORTED_MODULE_3__["tableTextFontFamily"],
        fontSize: _StylingConstants__WEBPACK_IMPORTED_MODULE_3__["tableTextFontSize"],
        lineHeight: _StylingConstants__WEBPACK_IMPORTED_MODULE_3__["tableTextLineHeight"],
        fontWeight: _StylingConstants__WEBPACK_IMPORTED_MODULE_3__["tableTextFontWeight"],
        color: _StylingConstants__WEBPACK_IMPORTED_MODULE_3__["colorLabels"].cellText,
        backgroundColor: 'inherit',
        overflow: 'hidden',
        resize: 'none',
        verticalAlign: 'middle',
        selectors: {
            ':focus': {
                outline: 'none'
            },
            [office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_2__["HighContrastSelector"]]: {
                background: 'window',
                color: 'windowText',
                MsHighContrastAdjust: 'none'
            }
        },
        textAlign: 'left',
        overflowWrap: 'break-word'
    }, inputMode === InputMode.invisibleOvertype && {
        selectors: {
            '::selection': {
                // We don't show the selection background color when in invisible overtype mode.
                background: 'transparent'
            }
        }
    }, styleOverrides);
    const onValueChange = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((ev) => {
        // tslint:disable-next-line: no-floating-promises
        dispatch({ type: 'setInputValue', inputValue: ev.target.value });
        // Update the input height here instead of in a React hook on the inputValue changing
        // to prevent flickering from re-rending multiple times
        updateTextAreaHeight();
    }, []);
    const onKeyDown = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((ev) => {
        // Remove focus from the input on escape without committing the value. Put focus where it was before the input took focus.
        // Also, restore the previous input value in the input box.
        if (ev.keyCode === office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["KeyCodes"].escape) {
            // tslint:disable-next-line: no-floating-promises
            dispatch({ type: 'setInputValue', inputValue: committedValue });
            // tslint:disable-next-line: no-floating-promises
            dispatch({ type: 'setInputMode', inputMode: InputMode.invisibleOvertype });
            if (inputMode === InputMode.ipEditMode) {
                ev.preventDefault();
            }
            return;
        }
        if (ev.keyCode === office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["KeyCodes"].enter) {
            ev.preventDefault();
            commitValueIfChanged(inputValue);
            return;
        }
        if (enableUndoRedo) {
            if (ev.ctrlKey && (ev.keyCode === office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["KeyCodes"].z || ev.keyCode === office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["KeyCodes"].y)) {
                // disable browser's default undo/redo.
                ev.preventDefault();
                return;
            }
        }
        // Keyboard shortcut for column menu, should not trigger EditMode
        if (ev.keyCode === office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["KeyCodes"].down && ev.altKey) {
            return;
        }
        // When in Edit Mode and want to exit via arrow keys.
        // Update: Removing check of edit mode as there are some
        // inconsistencies in current implementation of click.
        // Check whether caret was at the left most corner of cell when left key was pressed.
        if (ev.keyCode === office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["KeyCodes"].left && Object(_utils_Helper__WEBPACK_IMPORTED_MODULE_6__["isCaretAtStart"])(textAreaRef.current)) {
            callOnMoveSelectionFromCell(ev);
            return;
        }
        // Check whether caret was at the right most corner of cell when right key was pressed.
        if (ev.keyCode === office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["KeyCodes"].right && Object(_utils_Helper__WEBPACK_IMPORTED_MODULE_6__["isCaretAtEnd"])(textAreaRef.current)) {
            callOnMoveSelectionFromCell(ev);
            return;
        }
        // Check if up key was pressed.
        if (ev.keyCode === office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["KeyCodes"].up) {
            callOnMoveSelectionFromCell(ev);
            return;
        }
        // Check if down key was pressed.
        if (ev.keyCode === office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["KeyCodes"].down) {
            callOnMoveSelectionFromCell(ev);
            return;
        }
        // When the user is explicitly in IP edit mode, we want all the key down events to not bubble up. This way they can be handled
        // natively in the text area (ex: up and down arrow keys). However, any key down event that take the focus out of the cell
        // we still want to be handled by tablero itself.
        if (inputMode === InputMode.ipEditMode && !keyCodesThatRemoveFocus.has(ev.keyCode)) {
            ev.stopPropagation();
            return;
        }
        if (ev.shiftKey && ev.keyCode === office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["KeyCodes"].f10) {
            // TODO: 3585037: Implement Shift + F10 invoking the commanding surface.
        }
        // When F2 is pressed in overtype mode, we switch to IP edit mode and put the cursor at the end of the line.
        if (ev.keyCode === office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["KeyCodes"].f2 && inputMode === InputMode.invisibleOvertype && textAreaRef.current) {
            // tslint:disable-next-line: no-floating-promises
            dispatch({ type: 'setInputMode', inputMode: InputMode.ipEditMode });
            textAreaRef.current.setSelectionRange(inputValue.length, inputValue.length);
            return;
        }
        // When we tab to a new cell we start in invisible overtype. If we are typing we want to make sure we are in edit mode
        // so we can ensure we keep the appropriate events from bubbling up. However, we don't want this to be triggered on certain
        // keys that would move selection.
        if (inputMode === InputMode.invisibleOvertype && shouldEnterEditModeOnKeyDown(ev.keyCode)) {
            // tslint:disable-next-line: no-floating-promises
            dispatch({ type: 'setInputMode', inputMode: InputMode.ipEditMode });
        }
    }, [committedValue, inputMode, commitValueIfChanged, inputValue]);
    const shouldEnterEditModeOnKeyDown = (keyCode) => {
        // The left and right arrow keys should move selection. Not trigger edit mode.
        switch (keyCode) {
            case office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["KeyCodes"].left:
            case office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["KeyCodes"].right:
            case office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["KeyCodes"].up:
            case office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["KeyCodes"].down:
                return false;
            default:
                return true;
        }
    };
    const onTextAreaBlur = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"](() => {
        // Only update the aria-label dynamically if there is an additionalA11yLabel
        // e.g. For column headers, dynamically update to "Header cell: Task Title"
        // For body cells, a dynamic aria-label is not needed
        if (additionalA11yLabel) {
            Object(_CellHelpers__WEBPACK_IMPORTED_MODULE_7__["setAriaLabel"])(textAreaRef.current, `${inputValue}`);
        }
        commitValueIfChanged(inputValue);
        // tslint:disable-next-line: no-floating-promises
        dispatch({ type: 'setInputMode', inputMode: InputMode.unfocused });
    }, [inputValue, commitValueIfChanged, textAreaRef, additionalA11yLabel]);
    const onTextAreaMouseDown = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((ev) => {
        if (inputMode === InputMode.unfocused) {
            // Prevent the browser from setting the IP on the input box on first click
            ev.preventDefault();
        }
    }, [inputMode]);
    const onTextAreaClick = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"](() => {
        // We update the input mode on click instead of mouse down to make sure this happens after the browser has set the selection.
        if (inputMode === InputMode.unfocused) {
            // The first click of the input element should put it into overtype mode, where typing replaces all the current text.
            // tslint:disable-next-line: no-floating-promises
            dispatch({ type: 'setInputMode', inputMode: InputMode.invisibleOvertype });
        }
        else if (inputMode === InputMode.invisibleOvertype) {
            // A click once already inside of overtype mode should put it in ip edit mode.
            // tslint:disable-next-line: no-floating-promises
            dispatch({ type: 'setInputMode', inputMode: InputMode.ipEditMode });
        }
    }, [inputMode]);
    const onReplaceWhenInsertedComponentClosed = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((content) => {
        // tslint:disable-next-line: no-floating-promises
        dispatch({ type: 'setInputValue', inputValue: content });
    }, []);
    const onTextAreaInput = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((ev) => {
        // Pass input through intellisense to see if it matches any actions.
        //
        // We do this onInput instead of binding to when the inputValue changes to prevent recursive rendering of components being triggered.
        // Similarly calling this from onChange would do the same. If this was done onKeyDown the key would be passed to the new components being
        // rendered without they themselves handling the event which we want to keep handled inside Tablero.
        if (rowId && columnId && columnType && componentContext && onCellValueProposal && sendInsertComponentRequestData) {
            const boundingRect = textAreaRef.current && textAreaRef.current.getBoundingClientRect();
            const coordinates = boundingRect
                ? {
                    x: boundingRect.left,
                    y: boundingRect.top
                }
                : undefined;
            const ownerDocument = (textAreaRef.current && textAreaRef.current.ownerDocument) || undefined;
            const insertComponentRequestContext = {
                cellContext: {
                    rowId,
                    columnId,
                    columnType,
                    onCellValueProposal,
                    hostingRef: textAreaRef,
                    onReplaceWhenClosed: onReplaceWhenInsertedComponentClosed,
                    onMoveFocusToHostingCanvas: onUpdateCellSelection
                },
                hostContext: {
                    userInput: ev.target.value,
                    floatingUXContext: {
                        coordinates,
                        ownerDocument
                    },
                    // Tablero allows UX for inserting only 'inline' layout components for now in Production
                    // However, there is no technical limitation for hosting any other component.
                    canHandleInlineComponentsOnly: !Object(_table_TableAppFeatures__WEBPACK_IMPORTED_MODULE_8__["nestAnyLayoutFeatureEnabled"])()
                },
                componentContext,
                logger,
                sendInsertComponentRequestData
            };
            Object(_intelligence_InsertComponentRequest__WEBPACK_IMPORTED_MODULE_4__["notifyInsertComponentRequest"])(insertComponentRequestContext);
        }
    }, []);
    // If we received new committed value, replace the value in the input box with the new committed value.
    // TODO: TABLERO: We may not want to do this while the user is editing the cell, so we may need to track if the cell is in edit mode.
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        // tslint:disable-next-line: no-floating-promises
        dispatch({ type: 'setInputValue', inputValue: committedValue });
    }, [committedValue]);
    // Update the value of the text area shown to reflect the new input value
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        if (textAreaRef.current) {
            textAreaRef.current.value = inputValue;
            // For making it accessible for Read out reader
            textAreaRef.current.innerText = inputValue;
        }
        if (divRef.current) {
            divRef.current.innerText = inputValue;
            divRef.current.setAttribute('aria-label', (additionalA11yLabel !== undefined ? additionalA11yLabel : '') + ' ' + inputValue);
        }
    }, [inputValue, readOnly]);
    // When the selection state of the wrapping cell changes, be sure to update the input mode
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        var _a;
        // tslint:disable-next-line: no-floating-promises
        if (hasLocalSelection) {
            // Set into edit mode if we have moveSelectionToCellDetails but also make sure it is not readonly.
            if (moveSelectionToCellDetails && !((_a = textAreaRef.current) === null || _a === void 0 ? void 0 : _a.readOnly)) {
                dispatch({
                    type: 'setInputMode',
                    inputMode: InputMode.ipEditMode
                });
                onSelectionMovedToCell(moveSelectionToCellDetails);
            }
            else {
                dispatch({
                    type: 'setInputMode',
                    inputMode: InputMode.invisibleOvertype
                });
            }
            if (divRef.current) {
                divRef.current.focus();
            }
        }
        else {
            dispatch({
                type: 'setInputMode',
                inputMode: InputMode.unfocused
            });
        }
    }, [hasLocalSelection]);
    // If we enter invisible overtype mode for the first time, make sure we actually select all the text.
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        if (textAreaRef.current) {
            if (inputMode === InputMode.invisibleOvertype) {
                textAreaRef.current.select();
            }
        }
    }, [inputMode]);
    // On rendering the cell resize the height of the text area to support multi-line text
    // We can't do a simple hook to match componentDidMount (having [] as the dependencies) because on
    // initial render the input value isn't set and so we wouldn't get an accurate resize.
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        if (inputValue === committedValue && inputValue !== '') {
            updateTextAreaHeight();
        }
    }, [committedValue, inputValue]);
    // When the column resizes horizontally trigger an update to the text area height to account for
    // the changing div shape. This will only fire when the columnWidth prop has been changed.
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        if (columnWidth && inputValue !== '') {
            updateTextAreaHeight();
        }
    }, [columnWidth]);
    const onTextAreaFocus = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"](() => {
        // Only update the aria-label dynamically if there is an additionalA11yLabel
        // e.g. For column headers, dynamically update to "Header cell: Task Title"
        // For body cells, a dynamic aria-label is not needed
        if (additionalA11yLabel) {
            Object(_CellHelpers__WEBPACK_IMPORTED_MODULE_7__["setAriaLabel"])(textAreaRef.current, `${additionalA11yLabel} ${inputValue}`);
        }
    }, [textAreaRef, inputValue, additionalA11yLabel]);
    const onDivFocus = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"](() => {
        // Only update the aria-label dynamically if there is an additionalA11yLabel
        // e.g. For column headers, dynamically update to "Header cell: Task Title"
        // For body cells, a dynamic aria-label is not needed
        if (additionalA11yLabel) {
            Object(_CellHelpers__WEBPACK_IMPORTED_MODULE_7__["setAriaLabel"])(divRef.current, `${additionalA11yLabel} ${inputValue}`);
        }
    }, [divRef, inputValue, additionalA11yLabel]);
    const onDivBlur = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"](() => {
        // Only update the aria-label dynamically if there is an additionalA11yLabel
        // e.g. For column headers, dynamically update to "Header cell: Task Title"
        // For body cells, a dynamic aria-label is not needed
        if (additionalA11yLabel) {
            Object(_CellHelpers__WEBPACK_IMPORTED_MODULE_7__["setAriaLabel"])(divRef.current, `${inputValue}`);
        }
    }, [divRef, inputValue, additionalA11yLabel]);
    // TODO: Task 3720153: Make the cell uneditable when the row it's in is completed
    return readOnly ? (
    // Bug 4142596. Windows screen reader announces textarea with
    // word 'edit', even if it is marked readonly / aria-readonly.
    // For this reason, div is used for readonly cells.
    react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { className: cellClassName, ref: divRef, tabIndex: -1, onFocus: onDivFocus, onBlur: onDivBlur })) : (react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("textarea", { onInput: onTextAreaInput, onChange: onValueChange, className: cellClassName, onKeyDown: onKeyDown, onBlur: onTextAreaBlur, onFocus: onTextAreaFocus, onMouseDown: onTextAreaMouseDown, onClick: onTextAreaClick, rows: 1, ref: textAreaRef, onDrop: onDragAndDrop, onDragStart: onDragAndDrop }));
});


/***/ }),

/***/ "../tablero/lib/view/cells/CheckboxCell.js":
/*!*************************************************!*\
  !*** ../tablero/lib/view/cells/CheckboxCell.js ***!
  \*************************************************/
/*! exports provided: CheckboxCell */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "CheckboxCell", function() { return CheckboxCell; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/styling/7.13.5_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/styling/lib/index.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/memoize.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/setFocusVisibility.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/office-ui-fabric-react/7.121.6_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/office-ui-fabric-react/lib/components/Checkbox/Checkbox.js");
/* harmony import */ var _dataModel_adapterComponents_TableroAdapterComponentFactory__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../../dataModel/adapterComponents/TableroAdapterComponentFactory */ "../tablero/lib/dataModel/adapterComponents/TableroAdapterComponentFactory.js");
/* harmony import */ var _StylingConstants__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../StylingConstants */ "../tablero/lib/view/StylingConstants.js");
/* harmony import */ var _fluidx_icons__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! @fluidx/icons */ "../icons/lib/useIconNames/index.js");





const checkboxWrapperClassName = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["mergeStyles"])({
    display: 'flex',
    flexDirection: 'row',
    justifyContent: 'center',
    marginTop: '2px'
});
const checkmarkIconProps = {
    iconName: Object(_fluidx_icons__WEBPACK_IMPORTED_MODULE_7__["useComponentsIcon"])('FFXCCheckMark')
};
const checkboxStyles = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_2__["memoizeFunction"])((props) => {
    const { reversed, checked, disabled, indeterminate } = props;
    const checkboxColorChecked = _StylingConstants__WEBPACK_IMPORTED_MODULE_6__["colorLabels"].checkbox.colorChecked;
    const classNames = {
        checkbox: 'ms-Checkbox-checkbox',
        checkmark: 'ms-Checkbox-checkmark'
    };
    return {
        // office-ui-fabric-react is creating a border around Checkboxes when they have focus.
        // Since we don't want to display that border, we are overriding those styles.
        input: {
            selectors: {
                [`.${office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["IsFocusVisibleClassName"]} &:focus + label::before`]: {
                    outline: 'none'
                }
            }
        },
        checkbox: [
            {
                borderRadius: '50%',
                width: '16px',
                height: '16px',
                fontSize: '11px',
                selectors: {
                    [office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]]: {
                        borderColor: 'windowText'
                    }
                }
            },
            !disabled &&
                !indeterminate &&
                checked && {
                background: checkboxColorChecked,
                borderColor: checkboxColorChecked
            }
        ],
        label: {
            width: '16px'
        },
        root: [
            reversed && 'reversed',
            checked && 'is-checked',
            !disabled && 'is-enabled',
            disabled && 'is-disabled',
            !disabled && [
                !checked && {
                    selectors: {
                        [`:hover .${classNames.checkbox}`]: {
                            borderColor: checkboxColorChecked,
                            selectors: {
                                [office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]]: {
                                    borderColor: 'windowText'
                                }
                            }
                        },
                        [`:focus .${classNames.checkbox}`]: {
                            borderColor: checkboxColorChecked
                        },
                        [`:hover .${classNames.checkmark}`]: {
                            color: checkboxColorChecked,
                            selectors: {
                                [office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]]: {
                                    color: 'windowText'
                                }
                            }
                        }
                    }
                }
            ],
            checked &&
                !indeterminate && {
                selectors: {
                    [`:hover .${classNames.checkbox}`]: {
                        background: checkboxColorChecked,
                        borderColor: checkboxColorChecked,
                        selectors: {
                            [office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]]: {
                                borderColor: 'windowText',
                                background: 'windowText'
                            }
                        }
                    },
                    [`:focus .${classNames.checkbox}`]: {
                        background: checkboxColorChecked,
                        borderColor: checkboxColorChecked
                    },
                    [`.${classNames.checkbox}`]: {
                        background: checkboxColorChecked,
                        borderColor: checkboxColorChecked,
                        selectors: {
                            [office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]]: {
                                borderColor: 'windowText',
                                background: 'windowText'
                            }
                        }
                    },
                    [`:hover .${classNames.checkmark}, .${classNames.checkmark}`]: {
                        selectors: {
                            [office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]]: {
                                color: 'window'
                            }
                        }
                    }
                }
            }
        ]
    };
});
const CheckboxCell = react__WEBPACK_IMPORTED_MODULE_0__["memo"](function CheckboxCell(props) {
    const { rowId, columnId, columnType, value, onCellValueProposal, hasLocalSelection, onRowCompletionProposal, isReadOnlyMode } = props;
    const checkboxRef = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](null);
    const onCheckChanged = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((_, checked) => {
        // The fabric react component returns the new (desired) checked value for the checked parameter
        const newCellData = Object(_dataModel_adapterComponents_TableroAdapterComponentFactory__WEBPACK_IMPORTED_MODULE_5__["createAdapterComponent"])(columnType, !!checked);
        onCellValueProposal(rowId, columnId, newCellData);
    }, [rowId, columnId, onCellValueProposal]);
    // Put focus on the cell when the wrapping cell is selected
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        if (hasLocalSelection && checkboxRef.current) {
            checkboxRef.current.focus();
        }
    }, [hasLocalSelection]);
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        onRowCompletionProposal(value);
    }, [value]);
    return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { className: checkboxWrapperClassName },
        react__WEBPACK_IMPORTED_MODULE_0__["createElement"](office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_4__["Checkbox"], { checkmarkIconProps: checkmarkIconProps, checked: value, onChange: onCheckChanged, componentRef: checkboxRef, styles: checkboxStyles, disabled: isReadOnlyMode })));
});


/***/ }),

/***/ "../tablero/lib/view/cells/DateCell.js":
/*!*********************************************!*\
  !*** ../tablero/lib/view/cells/DateCell.js ***!
  \*********************************************/
/*! exports provided: DateCell */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "DateCell", function() { return DateCell; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _table_Table_props__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../table/Table.props */ "../tablero/lib/view/table/Table.props.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/styling/7.13.5_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/styling/lib/index.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/KeyCodes.js");
/* harmony import */ var _utilities_createComponent__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../../utilities/createComponent */ "../tablero/lib/utilities/createComponent.js");




const hostingDivClassName = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_2__["mergeStyles"])({
    textAlign: 'left',
    // We should not have an outline on the hosting div.
    // This makes sure the outline is not inherited from the host.
    outline: 'none'
});
const DateCell = react__WEBPACK_IMPORTED_MODULE_0__["memo"](function DateCell(props) {
    const { rowId, columnId, columnType, hasLocalSelection, onCellValueProposal, cellWrapperFocused, notifyCellWrapperFocusExecuted, componentContext, isReadOnlyMode } = props;
    const hostingRef = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](null);
    const loadDateComponent = () => {
        // tslint:disable-next-line: no-floating-promises
        Object(_utilities_createComponent__WEBPACK_IMPORTED_MODULE_4__["createComponentWithDataProvider"])(_table_Table_props__WEBPACK_IMPORTED_MODULE_1__["DataTypeMap"][columnType], componentContext.containerRuntime, {
            isTextBoxMode: true
        }).then((result) => {
            if (result) {
                onCellValueProposal(rowId, columnId, result);
            }
        });
    };
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        if (hasLocalSelection && !isReadOnlyMode) {
            loadDateComponent();
        }
    }, [hasLocalSelection]);
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        if (cellWrapperFocused) {
            notifyCellWrapperFocusExecuted();
        }
    }, [cellWrapperFocused]);
    const onKeyDownHostingDiv = (ev) => {
        if (ev.altKey && ev.keyCode === office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["KeyCodes"].down) {
            // Make sure we don't propagate the event to tablero to change the cell selection.
            ev.stopPropagation();
        }
        if (ev.keyCode === office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["KeyCodes"].escape && hostingRef.current) {
            hostingRef.current.focus();
        }
    };
    return react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { ref: hostingRef, className: hostingDivClassName, onKeyDown: onKeyDownHostingDiv, tabIndex: 0 });
});


/***/ }),

/***/ "../tablero/lib/view/cells/FluidComponentCell.js":
/*!*******************************************************!*\
  !*** ../tablero/lib/view/cells/FluidComponentCell.js ***!
  \*******************************************************/
/*! exports provided: FluidComponentCell */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "FluidComponentCell", function() { return FluidComponentCell; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _table_Table_props__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../table/Table.props */ "../tablero/lib/view/table/Table.props.js");
/* harmony import */ var _dataModel_adapterComponents_TableroAdapterComponentFactory__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../../dataModel/adapterComponents/TableroAdapterComponentFactory */ "../tablero/lib/dataModel/adapterComponents/TableroAdapterComponentFactory.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/KeyCodes.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/styling/7.13.5_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/styling/lib/index.js");
/* harmony import */ var _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! @fluidx/office-fluid-types */ "../office-fluid-types/lib/ComponentContracts/ComponentRichTextEditor.js");
/* harmony import */ var _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! @fluidx/office-fluid-types */ "../office-fluid-types/lib/ComponentContracts/ComponentSelection.js");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../../telemetry */ "../tablero/lib/telemetry/ErrorEvents.js");
/* harmony import */ var _ms_office_bohemia_ux__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! @ms/office-bohemia-ux */ "../office-bohemia-ux/lib/eventing/matchEventWithOrigin.js");
/* harmony import */ var _ms_office_bohemia_ux__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! @ms/office-bohemia-ux */ "../office-bohemia-ux/lib/eventing/CommandEvent.js");
/* harmony import */ var _fluidx_utilities__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! @fluidx/utilities */ "../utilities/lib/logger/LoggerHelpers.js");
/* harmony import */ var _table_TableKeyboarding__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(/*! ../table/TableKeyboarding */ "../tablero/lib/view/table/TableKeyboarding.js");









const allowedKeysInReadOnlyMode = new Set([
    // Prevent deletion in read only mode.
    office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["KeyCodes"].left,
    office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["KeyCodes"].right,
    office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["KeyCodes"].up,
    office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["KeyCodes"].down,
    office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["KeyCodes"].pageUp,
    office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["KeyCodes"].pageDown,
    office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["KeyCodes"].escape,
    office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["KeyCodes"].tab
]);
const isSafari = navigator.userAgent.indexOf('Chrome') === -1 && navigator.userAgent.indexOf('Safari') >= 0;
const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent) && !window.MSStream;
const hostingDivClassName = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_4__["mergeStyles"])({
    width: '100%',
    overflow: 'hidden',
    whiteSpace: 'nowrap',
    textAlign: 'left',
    selectors: {
        ':focus': {
            outline: 'none'
        }
    },
    backgroundColor: 'inherit'
}, 
// There is an issue the hosting element does not change its height along with content increase in webkit browsers (non-chromium).
// This happens in Safari as well as all browsers in IOS devices (since IOS browsers are all based on webkit), which caused the last line
// of content is hidden to the user.
// TODO: After this fix, there is still another issue for webkit browsers - when row height increases, the hosting element
// doesn't expand accordingly to take up the entire the cell space.
isSafari || isIOS ? { height: 'auto', minHeight: '100%' } : { height: '100%' });
const FluidComponentCell = react__WEBPACK_IMPORTED_MODULE_0__["memo"]((props) => {
    const { rowId, columnId, columnType, componentUrl, onCellValueProposal, hasLocalSelection, uiContextComponentUrl, onUpdateCellSelection, componentContext, logger, onMoveSelectionFromCell, onNestedComponentDataChangeCallback, isReadOnlyMode, isRowCompleted } = props;
    // Ref to the element which is hosting the component.
    const hostingDivRef = react__WEBPACK_IMPORTED_MODULE_0__["createRef"]();
    const nestedComponent = react__WEBPACK_IMPORTED_MODULE_0__["useRef"]();
    const isHostedComponentFocused = () => {
        var _a;
        return (((_a = nestedComponent === null || nestedComponent === void 0 ? void 0 : nestedComponent.current) === null || _a === void 0 ? void 0 : _a.ComponentFocusable) !== undefined &&
            nestedComponent.current.ComponentFocusable.isComponentFocused());
    };
    const registerSelectionChangeCallback = () => {
        var _a;
        const componentSelection = (_a = nestedComponent.current) === null || _a === void 0 ? void 0 : _a.ComponentSelection;
        if (componentSelection === null || componentSelection === void 0 ? void 0 : componentSelection.setHostComponentSelection) {
            componentSelection.setHostComponentSelection(getCellComponentSelection());
        }
    };
    const registerDataChangeCallback = () => {
        var _a;
        const subComponent = (_a = nestedComponent.current) === null || _a === void 0 ? void 0 : _a.ComponentProduceString;
        if (subComponent !== undefined && subComponent.registerDataChangeCallback) {
            subComponent.registerDataChangeCallback(onNestedComponentDataChangeCallback);
        }
    };
    const callOnMoveSelectionFromCell = (ev) => {
        onMoveSelectionFromCell({
            type: 'cell',
            columnId,
            rowId
        })(ev);
    };
    const getCellComponentSelection = () => {
        const fluidCellComponentSelection = {
            get ComponentSelection() {
                return this;
            },
            setSelection: (selectionParams) => {
                if (selectionParams.direction !== undefined) {
                    callOnMoveSelectionFromCell({ keyCode: _table_TableKeyboarding__WEBPACK_IMPORTED_MODULE_11__["getKeyCodeFromSelectionDirections"][selectionParams.direction] });
                    return true;
                }
                return false;
            }
        };
        return fluidCellComponentSelection;
    };
    const unregisterDataChangeCallback = () => {
        var _a;
        const subComponent = (_a = nestedComponent.current) === null || _a === void 0 ? void 0 : _a.ComponentProduceString;
        if (subComponent !== undefined && subComponent.unregisterDataChangeCallback) {
            subComponent.unregisterDataChangeCallback(onNestedComponentDataChangeCallback);
        }
    };
    const registerKeyBindingHandlers = () => {
        var _a, _b;
        if ((_b = (_a = nestedComponent === null || nestedComponent === void 0 ? void 0 : nestedComponent.current) === null || _a === void 0 ? void 0 : _a.ComponentRichTextEditor) === null || _b === void 0 ? void 0 : _b.registerKeyBindingHandlers) {
            const hostingDiv = hostingDivRef.current;
            nestedComponent === null || nestedComponent === void 0 ? void 0 : nestedComponent.current.ComponentRichTextEditor.registerKeyBindingHandlers([
                {
                    action: {
                        query: () => {
                            return true;
                        },
                        execute: () => {
                            callOnMoveSelectionFromCell({ keyCode: office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["KeyCodes"].tab });
                        }
                    },
                    commonKeyCombinations: [{ keyCode: office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["KeyCodes"].tab }]
                },
                {
                    action: {
                        query: () => {
                            return true;
                        },
                        execute: () => {
                            callOnMoveSelectionFromCell({ keyCode: office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["KeyCodes"].tab, shiftKey: true });
                        }
                    },
                    commonKeyCombinations: [{ keyCode: office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["KeyCodes"].tab, shiftKey: true }]
                },
                {
                    action: {
                        query: () => {
                            return true;
                        },
                        execute: () => {
                            callOnMoveSelectionFromCell({ keyCode: office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["KeyCodes"].enter });
                        }
                    },
                    commonKeyCombinations: [{ keyCode: office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["KeyCodes"].enter }]
                },
                {
                    action: {
                        query: () => {
                            return true;
                        },
                        execute: () => {
                            hostingDiv === null || hostingDiv === void 0 ? void 0 : hostingDiv.focus();
                        }
                    },
                    commonKeyCombinations: [{ keyCode: office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["KeyCodes"].escape }]
                }
            ]);
        }
    };
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        if (hasLocalSelection) {
            // If the cell has local selection and the component is focused then do not focus the hosting div
            const hostingDiv = hostingDivRef.current;
            if (hostingDiv && !isHostedComponentFocused()) {
                hostingDiv.focus();
            }
        }
    }, [hasLocalSelection]);
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        var _a;
        const componentRichTextEditor = (_a = nestedComponent.current) === null || _a === void 0 ? void 0 : _a.ComponentRichTextEditor;
        if (columnType === _table_Table_props__WEBPACK_IMPORTED_MODULE_1__["DataType"].SimpleRichText && componentRichTextEditor && componentRichTextEditor.setTextFormat) {
            componentRichTextEditor.setTextFormat({
                [_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_5__["ComponentRichTextFormat"].strikethrough]: isRowCompleted
            });
        }
    }, [isRowCompleted]);
    // TODO: Task 3338117: Resize hosted component to fit inside hosting div
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        const hostingDiv = hostingDivRef.current;
        if (hostingDiv && componentUrl) {
            Object(_ms_office_bohemia_ux__WEBPACK_IMPORTED_MODULE_8__["markAsHostingElement"])(hostingDiv);
            hostingDiv.addEventListener(_ms_office_bohemia_ux__WEBPACK_IMPORTED_MODULE_9__["commandEventName"], (event) => {
                const commandEvent = event;
                if (commandEvent) {
                    switch (commandEvent.commandName) {
                        case _ms_office_bohemia_ux__WEBPACK_IMPORTED_MODULE_9__["CustomEventName"].replaceNodeWithContent:
                            deleteComponentReference(commandEvent.commandValue);
                            event.stopPropagation();
                            break;
                        case _ms_office_bohemia_ux__WEBPACK_IMPORTED_MODULE_9__["CustomEventName"].moveFocusToHostingCanvas: {
                            hostingDiv.focus();
                            onUpdateCellSelection();
                            event.stopPropagation();
                            break;
                        }
                    }
                }
            });
            loadComponent(componentUrl)
                .then((component) => {
                var _a;
                if (!component) {
                    // TODO: we could use a N/A component experience designed
                    logger &&
                        Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_10__["sendErrorEvent"])(logger, {
                            eventName: _telemetry__WEBPACK_IMPORTED_MODULE_7__["ErrorEvent"].ComponentNestingError,
                            errorCode: _telemetry__WEBPACK_IMPORTED_MODULE_7__["ErrorCode"].ComponentNestingError.NestFailedToLoad
                        });
                    return;
                }
                // Keep reference to loaded component to be able to focus after
                nestedComponent.current = component;
                // Register Callback function so that whenever nested component can't handle
                // the cursor navigation anymore, it passes control back to Table.
                registerSelectionChangeCallback();
                registerDataChangeCallback();
                const visual = component.IComponentHTMLVisual;
                const uiContextProvider = {
                    IUIContextProvider: {
                        getUIContext: (_triggerComponent) => {
                            return { uiContextComponentUrl, deepPath: '' };
                        }
                    }
                };
                // Pass uiContextProvider to child components that do have addView. Child components can use the uiContextProvider
                // to discover how to reproduce the view they are a part of (e.g. for @mention notification emails)
                // Note: addView accepts an IComponent, i.e. more than just uiContextProvider can be passed in
                // BACKCOMPAT - For backcompat, several components provide a .IComponentHTMLVisual (but no addView).
                // This check for .addView existence is not necessary after that backcompat is removed.
                // const componentRenderer = visual ? visual.addView(uiContextProvider) : component.IComponentHTMLView;
                // TODO: Remove when we believe OWH < 0.14 is no longer in use.
                const componentRenderer = visual
                    ? visual.addView
                        ? visual.addView(uiContextProvider)
                        : component.IComponentHTMLView
                            ? component.IComponentHTMLView
                            : // An IComponentHTMLVisual without an .addView is going to be the old version, which has a .render.
                                visual
                    : component.IComponentHTMLView;
                if ((_a = component.ComponentFocusable) === null || _a === void 0 ? void 0 : _a.setHostComponentFocusable) {
                    component.ComponentFocusable.setHostComponentFocusable(getHostCellComponentFocusable());
                }
                if (componentRenderer) {
                    componentRenderer.render(hostingDiv, {
                        display: 'block',
                        maxWidth: Number.MAX_VALUE,
                        paddingTop: 0,
                        paddingBottom: 0,
                        paddingLeft: 0,
                        paddingRight: 0,
                        role: 'cell'
                    });
                }
                else {
                    logger &&
                        Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_10__["sendTelemetryEvent"])(logger, {
                            eventName: _telemetry__WEBPACK_IMPORTED_MODULE_7__["ErrorEvent"].ComponentNestingError,
                            errorCode: _telemetry__WEBPACK_IMPORTED_MODULE_7__["ErrorCode"].ComponentNestingError.NestHasNoRenderableVisuals
                        });
                    // TODO: Render value of loadable component that didn't provide visual
                }
                registerKeyBindingHandlers();
                // If we are loading and rendering nested component when Table cell already has local selection,
                // then give focus to the nested component.
                // TODO: Need to define the behavior once we support navigation in View Mode for nested component.
                // In this case, focus should be at hosting div and not inside the nested component.
                if (hasLocalSelection) {
                    // Adding this job to the end of queue for pending asynchronous tasks in the queue to finish
                    // after component render.
                    setTimeout(() => {
                        var _a, _b;
                        (_b = (_a = nestedComponent.current) === null || _a === void 0 ? void 0 : _a.ComponentFocusable) === null || _b === void 0 ? void 0 : _b.giveFocus();
                    }, 0);
                }
            })
                .catch((error) => {
                // TODO: we could use a N/A component experience designed
                logger &&
                    Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_10__["sendErrorEvent"])(logger, {
                        eventName: _telemetry__WEBPACK_IMPORTED_MODULE_7__["ErrorEvent"].ComponentNestingError,
                        errorCode: _telemetry__WEBPACK_IMPORTED_MODULE_7__["ErrorCode"].ComponentNestingError.NestFailedToLoad
                    }, error);
            });
        }
    }, [componentUrl]);
    const getHostCellComponentFocusable = () => {
        // Return an implementation of componentFocusable that allows nested component to call giveFocus to return focus back to tablero cell
        const tableComponentFocusable = {
            get ComponentFocusable() {
                return this;
            },
            giveFocus: () => {
                if (hostingDivRef.current) {
                    hostingDivRef.current.focus();
                    onUpdateCellSelection();
                    return true;
                }
                return false;
            },
            isComponentFocused: () => {
                // TODO:Task 4245381: Pass Tablero isComponentFocused here
                return false;
            }
        };
        return tableComponentFocusable;
    };
    const deleteComponentReference = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((replaceWithValue) => {
        if (componentUrl) {
            unregisterDataChangeCallback();
            const newCellData = Object(_dataModel_adapterComponents_TableroAdapterComponentFactory__WEBPACK_IMPORTED_MODULE_2__["createAdapterComponent"])(columnType, replaceWithValue || '');
            onCellValueProposal(rowId, columnId, newCellData);
        }
    }, [componentUrl]);
    const loadComponent = async (componentIdentifier) => {
        const response = await componentContext.containerRuntime.request({ url: componentIdentifier });
        if (response.status !== 200 || response.mimeType !== 'fluid/component') {
            return;
        }
        return response.value;
    };
    const isEventFromHost = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((e) => {
        const composedPath = e.nativeEvent.composedPath ? e.nativeEvent.composedPath() : Object(_ms_office_bohemia_ux__WEBPACK_IMPORTED_MODULE_8__["buildEventPath"])(e.nativeEvent);
        return Object(_ms_office_bohemia_ux__WEBPACK_IMPORTED_MODULE_8__["isEventFromComponentHostingElement"])(e.currentTarget, composedPath);
    }, []);
    // TODO: Task 3417784: How do we handle keyboard events between nested components and tablero?
    // TODO: Task 3417788: What's the UI entry point for loading components?
    const handleKeyDown = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((e) => {
        var _a, _b, _c, _d;
        if (isEventFromHost(e)) {
            return;
        }
        // Restrict read only user to use only the navigation keys
        if (isReadOnlyMode && !allowedKeysInReadOnlyMode.has(e.keyCode)) {
            e.preventDefault();
            e.stopPropagation();
            return;
        }
        if (e.keyCode === office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["KeyCodes"].backspace || e.keyCode === office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["KeyCodes"].del) {
            deleteComponentReference();
        }
        else if (e.keyCode === office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["KeyCodes"].f2) {
            // Try to give focus to the component
            (_b = (_a = nestedComponent === null || nestedComponent === void 0 ? void 0 : nestedComponent.current) === null || _a === void 0 ? void 0 : _a.ComponentFocusable) === null || _b === void 0 ? void 0 : _b.giveFocus();
        }
        else if (e.keyCode === 32 || // space
            (e.keyCode > 47 && e.keyCode < 91) || // alphanumeric characters
            (e.keyCode > 95 && e.keyCode < 112) || // numpad keys
            (e.keyCode > 185 && e.keyCode < 193) || // ;=,-./`
            (e.keyCode > 218 && e.keyCode < 223) // [\]'
        ) {
            // TODO: Do we want to allow other characters ? Is there a better way to do this check for allowed keys?
            if (columnType === _table_Table_props__WEBPACK_IMPORTED_MODULE_1__["DataType"].Text) {
                // If we are not in a text column then give focus to the component and let it handle overtyping
                deleteComponentReference(e.key);
            }
            else {
                (_d = (_c = nestedComponent === null || nestedComponent === void 0 ? void 0 : nestedComponent.current) === null || _c === void 0 ? void 0 : _c.ComponentSelection) === null || _d === void 0 ? void 0 : _d.setSelection({ mode: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_6__["SelectionMode"].fullRange });
            }
        }
        else if (e.keyCode !== office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["KeyCodes"].escape) {
            // Do not prevent default in case of escape key as ComponentUX checks for isDefaultPrevented
            // to take away the focus.
            e.preventDefault();
        }
    }, [isReadOnlyMode]);
    return react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { ref: hostingDivRef, className: hostingDivClassName, onKeyDown: handleKeyDown, tabIndex: 0 });
});


/***/ }),

/***/ "../tablero/lib/view/cells/GrabberCell.js":
/*!************************************************!*\
  !*** ../tablero/lib/view/cells/GrabberCell.js ***!
  \************************************************/
/*! exports provided: GrabberCell */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "GrabberCell", function() { return GrabberCell; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _StylingConstants__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../StylingConstants */ "../tablero/lib/view/StylingConstants.js");
/* harmony import */ var _cells_CellDefinitions__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../cells/CellDefinitions */ "../tablero/lib/view/cells/CellDefinitions.js");
/* harmony import */ var _GrabberCellStyles__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./GrabberCellStyles */ "../tablero/lib/view/cells/GrabberCellStyles.js");
/* harmony import */ var _utils_ResizeObserverWrapper__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../utils/ResizeObserverWrapper */ "../tablero/lib/view/utils/ResizeObserverWrapper.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/office-ui-fabric-react/7.121.6_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/office-ui-fabric-react/lib/components/Callout/Callout.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/office-ui-fabric-react/7.121.6_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/office-ui-fabric-react/lib/common/DirectionalHint.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/DelayedRender.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/office-ui-fabric-react/7.121.6_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/office-ui-fabric-react/lib/components/Icon/FontIcon.js");
/* harmony import */ var _utils_Helper__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ../utils/Helper */ "../tablero/lib/view/utils/Helper.js");
/* harmony import */ var _fluidx_icons__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! @fluidx/icons */ "../icons/lib/useIconNames/index.js");
/* harmony import */ var _table_TableAppFeatures__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(/*! ../table/TableAppFeatures */ "../tablero/lib/view/table/TableAppFeatures.js");
/* harmony import */ var _StringTable__WEBPACK_IMPORTED_MODULE_12__ = __webpack_require__(/*! ../StringTable */ "../tablero/lib/view/StringTable.js");










// The GrabberContainer is lazily loaded to improve performance.
const GrabberContainer = react__WEBPACK_IMPORTED_MODULE_0__["lazy"](() => __webpack_require__.e(/*! import() | GrabberContainer */ "GrabberContainer").then(__webpack_require__.bind(null, /*! ../grabber/GrabberContainer */ "../tablero/lib/view/grabber/GrabberContainer.js")).then((module) => ({
    default: module.GrabberContainer
})));
/**
 * Grabber Cell is a <td> element that can be one of the following
 * a) It is empty in case no GrabberContainer is not to be shown.
 * b) Attaches GrabberContainer dynamically to DOM based upon different conditions.
 */
const GrabberCell = react__WEBPACK_IMPORTED_MODULE_0__["memo"](function GrabberCell(props) {
    const { getPreset, id, type, columnWidth, getLastKnownLocalPresence, onLocalPresenceValueProposal, onCellStyleOverride, onInsertProposal, onDeleteProposal, onSortByColumnProposal, sortDirection, hoveredTableCellId, rowHeight, tableWrapperElm, isInitialCell, isColumnInResizeMode, onResizeColumnLeaveCmd, isSelected, onSelect, isReadOnlyMode, onMouseDown, isLastCell } = props;
    const setGrabberCellSelected = (isSelected) => {
        isSelected ? onSelect(id) : onSelect(undefined);
    };
    const [isGrabberCellHovered, setGrabberCellHovered] = react__WEBPACK_IMPORTED_MODULE_0__["useState"](false);
    const [bounds, setBounds] = react__WEBPACK_IMPORTED_MODULE_0__["useState"](undefined);
    const onFocusCapture = (event) => {
        // Clear the local user's presence from the table cells
        onLocalPresenceValueProposal(undefined);
        // React focus event bubbles up, hence stop it
        event.stopPropagation();
    };
    /**
     * Callback function that is invoked when grabber is clicked
     */
    const onGrabberCellSelectProposal = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((isSelected) => {
        setGrabberCellSelected(isSelected);
        if (!isSelected) {
            // Restore the local user's presence in table cell when selection is dismissed
            onLocalPresenceValueProposal(getLastKnownLocalPresence());
        }
    }, [setGrabberCellSelected]);
    const grabberRef = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](null);
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        // Only run this effect when resize column from commanding menu feature is enabled.
        if (Object(_table_TableAppFeatures__WEBPACK_IMPORTED_MODULE_11__["resizeColumnFromCmdMenuFeatureEnabled"])()) {
            onGrabberCellSelectProposal(!!isColumnInResizeMode);
            onCellStyleOverride(isColumnInResizeMode
                ? {
                    allEdgesOverride: { boundaryColor: _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["colorLabels"].grabber.clickedBorder, columnId: id },
                    singleEdgeOverride: _cells_CellDefinitions__WEBPACK_IMPORTED_MODULE_2__["CellStyleOverrideOption"].Default
                }
                : {
                    allEdgesOverride: _cells_CellDefinitions__WEBPACK_IMPORTED_MODULE_2__["CellStyleOverrideOption"].Default
                });
            // Initialize bounds for resize column callout when in resize mode.
            // This also handles the case when the window is resized between two resize column invokes changing the tableWrapperElm right bound value.
            const tableWrapperElemClientRect = tableWrapperElm === null || tableWrapperElm === void 0 ? void 0 : tableWrapperElm.getBoundingClientRect();
            if (isColumnInResizeMode && tableWrapperElemClientRect) {
                setBounds(Object(_utils_Helper__WEBPACK_IMPORTED_MODULE_9__["getResizeColumnCalloutBounds"])(tableWrapperElemClientRect));
            }
            // Updating resize column callout bounds when:
            // 1. tableWrapperElm top bound changes due to horizontal scrollbar appearing/disappearing on increasing/decreasing column width.
            // 2. tableWrapperElm right bound changes due to increase/decrease in column width.
            let resizeObserver = undefined;
            // tslint:disable-next-line: no-floating-promises
            Object(_utils_ResizeObserverWrapper__WEBPACK_IMPORTED_MODULE_4__["ensureResizeObserver"])().then(() => {
                if (isColumnInResizeMode && tableWrapperElm && grabberRef && grabberRef.current) {
                    resizeObserver = new _utils_ResizeObserverWrapper__WEBPACK_IMPORTED_MODULE_4__["default"](() => {
                        setBounds(Object(_utils_Helper__WEBPACK_IMPORTED_MODULE_9__["getResizeColumnCalloutBounds"])(tableWrapperElm.getBoundingClientRect()));
                    });
                    resizeObserver.observe(grabberRef.current);
                }
            });
            const handleMouseWheel = () => {
                onResizeColumnLeaveCmd && onResizeColumnLeaveCmd();
            };
            // When the user intends to scroll in the document or in the table we exit the resize mode.
            if (isColumnInResizeMode) {
                document.addEventListener('wheel', handleMouseWheel, { capture: true, passive: true, once: true });
            }
            // Cleanup - removing resizeObserver and wheel event listener on exiting resize mode.
            return () => {
                // Remove all elements from observer
                if (resizeObserver)
                    resizeObserver.disconnect();
                // Fallback for any browser which does not support 'once' EventListenerOption.
                document.removeEventListener('wheel', handleMouseWheel, { capture: true });
            };
        }
        return;
    }, [isColumnInResizeMode]);
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        if (!isColumnInResizeMode && isSelected && grabberRef.current) {
            // Focus the current Grabber Cell if it is selected
            grabberRef.current.focus();
        }
    }, [isSelected, isColumnInResizeMode]);
    /**
     * Make the grabber visible on mouse enter.
     */
    const onMouseEnter = () => {
        setGrabberCellHovered(true);
    };
    /**
     * Hide the grabber on mouse leave.
     */
    const onMouseLeave = () => {
        setGrabberCellHovered(false);
    };
    /**
     * decides if Grabber Container needs to be attached dynamically or not,
     * if current grabber cell is hovered,
     * or corresponding row/column cell is hovered,
     * or grabber cell has been selected.
     */
    const shouldAttachGrabber = isGrabberCellHovered || hoveredTableCellId === id || isSelected;
    /**
     * Track if the right edge of the last cell is visible.
     */
    const [isRightEdgeVisible, setRightEdgeVisibility] = react__WEBPACK_IMPORTED_MODULE_0__["useState"](false);
    /**
     * Track if the left edge of the first cell (also termed as initial cell) is visible.
     */
    const [isLeftEdgeVisible, setLeftEdgeVisibility] = react__WEBPACK_IMPORTED_MODULE_0__["useState"](false);
    const grabberCellClassName = type === 0 /* Row */
        ? Object(_GrabberCellStyles__WEBPACK_IMPORTED_MODULE_3__["grabberRowClassName"])(isSelected, rowHeight, false /*isPreview*/)
        : Object(_GrabberCellStyles__WEBPACK_IMPORTED_MODULE_3__["grabberColumnClassName"])(isSelected, columnWidth, isInitialCell || isLastCell ? true : false, false /*isPreview*/, window.devicePixelRatio);
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        // Only run this for the initial and last cell
        if (isInitialCell || isLastCell) {
            // Tracks if an animation frame request has been queued to handle a scroll event.
            // Prevents negative perf implications of multiple handlers running per animation frame
            let isRafScheduled = false;
            const scrollHandler = (ev) => {
                // If an animation frame request is scheduled, do nothing
                if (!isRafScheduled) {
                    // Only execute this when the browser is about to paint
                    requestAnimationFrame(() => {
                        if (ev.target instanceof HTMLElement) {
                            // Update visibility of cell
                            const edgeVisibility = Object(_utils_Helper__WEBPACK_IMPORTED_MODULE_9__["getEdgeVisibilityOfElementInContainer"])(grabberRef === null || grabberRef === void 0 ? void 0 : grabberRef.current, ev.target);
                            setRightEdgeVisibility(edgeVisibility.isRightEdgeVisible);
                            setLeftEdgeVisibility(edgeVisibility.isLeftEdgeVisible);
                        }
                        // Animation frame request has been processed
                        isRafScheduled = false;
                    });
                    // Animation frame request has been scheduled
                    isRafScheduled = true;
                }
            };
            const edgeVisibility = Object(_utils_Helper__WEBPACK_IMPORTED_MODULE_9__["getEdgeVisibilityOfElementInContainer"])(grabberRef === null || grabberRef === void 0 ? void 0 : grabberRef.current, tableWrapperElm);
            setLeftEdgeVisibility(edgeVisibility.isLeftEdgeVisible);
            setRightEdgeVisibility(edgeVisibility.isRightEdgeVisible);
            tableWrapperElm === null || tableWrapperElm === void 0 ? void 0 : tableWrapperElm.addEventListener('scroll', scrollHandler);
            return () => {
                tableWrapperElm === null || tableWrapperElm === void 0 ? void 0 : tableWrapperElm.removeEventListener('scroll', scrollHandler);
            };
        }
        return;
    });
    return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("td", { ref: grabberRef, className: grabberCellClassName, onMouseEnter: onMouseEnter, onMouseLeave: onMouseLeave, 
        // Allows focus on Grabber cell when clicked
        tabIndex: -1, onFocusCapture: onFocusCapture },
        isColumnInResizeMode && (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_5__["Callout"], { className: _GrabberCellStyles__WEBPACK_IMPORTED_MODULE_3__["resizeColumnCalloutStyleSet"].callout, target: `.${grabberCellClassName}`, role: "status", "aria-live": "assertive", gapSpace: _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["resizeColumnCalloutStyleConstants"].gapSpace, directionalHint: office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_6__["DirectionalHint"].topCenter, beakWidth: _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["resizeColumnCalloutStyleConstants"].beakWidth, preventDismissOnScroll: true, bounds: bounds },
            react__WEBPACK_IMPORTED_MODULE_0__["createElement"](office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_7__["DelayedRender"], null,
                react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("span", { className: _GrabberCellStyles__WEBPACK_IMPORTED_MODULE_3__["resizeColumnCalloutStyleSet"].textHolder, "aria-label": `${_StringTable__WEBPACK_IMPORTED_MODULE_12__["resizeTextStart"]} ${_StringTable__WEBPACK_IMPORTED_MODULE_12__["resizeTextLeftArrow"]} ${_StringTable__WEBPACK_IMPORTED_MODULE_12__["resizeTextMiddle"]} ${_StringTable__WEBPACK_IMPORTED_MODULE_12__["resizeTextRightArrow"]} ${_StringTable__WEBPACK_IMPORTED_MODULE_12__["resizeTextEnd"]}` },
                    react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("p", { className: _GrabberCellStyles__WEBPACK_IMPORTED_MODULE_3__["resizeColumnCalloutStyleSet"].text, "aria-hidden": "true" }, _StringTable__WEBPACK_IMPORTED_MODULE_12__["resizeTextStart"]),
                    react__WEBPACK_IMPORTED_MODULE_0__["createElement"](office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_8__["FontIcon"], { iconName: Object(_fluidx_icons__WEBPACK_IMPORTED_MODULE_10__["useComponentsIcon"])('FFXCReplyAlt'), className: _GrabberCellStyles__WEBPACK_IMPORTED_MODULE_3__["resizeColumnCalloutStyleSet"].fontIconLeft }),
                    react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("p", { className: _GrabberCellStyles__WEBPACK_IMPORTED_MODULE_3__["resizeColumnCalloutStyleSet"].text, "aria-hidden": "true" }, _StringTable__WEBPACK_IMPORTED_MODULE_12__["resizeTextMiddle"]),
                    react__WEBPACK_IMPORTED_MODULE_0__["createElement"](office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_8__["FontIcon"], { iconName: Object(_fluidx_icons__WEBPACK_IMPORTED_MODULE_10__["useComponentsIcon"])('FFXCReplyAlt'), className: _GrabberCellStyles__WEBPACK_IMPORTED_MODULE_3__["resizeColumnCalloutStyleSet"].fontIconRight }),
                    react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("p", { className: _GrabberCellStyles__WEBPACK_IMPORTED_MODULE_3__["resizeColumnCalloutStyleSet"].text, "aria-hidden": "true" }, _StringTable__WEBPACK_IMPORTED_MODULE_12__["resizeTextEnd"]))))),
        shouldAttachGrabber && (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](react__WEBPACK_IMPORTED_MODULE_0__["Suspense"], { fallback: null },
            react__WEBPACK_IMPORTED_MODULE_0__["createElement"](GrabberContainer, { grabberRef: grabberRef, id: id, onInsertProposal: onInsertProposal, onDeleteProposal: onDeleteProposal, isGrabberCellSelected: isSelected, onGrabberCellSelectProposal: onGrabberCellSelectProposal, onCellStyleOverride: onCellStyleOverride, getPreset: getPreset, sortDirection: sortDirection, onSortByColumnProposal: onSortByColumnProposal, tableWrapperElm: tableWrapperElm, type: type, isInitialCell: isInitialCell, isColumnInResizeMode: isColumnInResizeMode, onMouseDown: onMouseDown, isLastCell: isLastCell, isReadOnlyMode: isReadOnlyMode, isLeftEdgeVisible: isLeftEdgeVisible, isRightEdgeVisible: isRightEdgeVisible })))));
});


/***/ }),

/***/ "../tablero/lib/view/cells/GrabberCellStyles.js":
/*!******************************************************!*\
  !*** ../tablero/lib/view/cells/GrabberCellStyles.js ***!
  \******************************************************/
/*! exports provided: grabberRowClassName, grabberColumnClassName, resizeColumnCalloutStyleSet */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "grabberRowClassName", function() { return grabberRowClassName; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "grabberColumnClassName", function() { return grabberColumnClassName; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "resizeColumnCalloutStyleSet", function() { return resizeColumnCalloutStyleSet; });
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/memoize.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/styling/7.13.5_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/styling/lib/index.js");
/* harmony import */ var _StylingConstants__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../StylingConstants */ "../tablero/lib/view/StylingConstants.js");
/* harmony import */ var _utils_getBoxShadow__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../utils/getBoxShadow */ "../tablero/lib/view/utils/getBoxShadow.js");



const grabberColorLabels = _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["colorLabels"].grabber;
const grabberCommonClassName = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["memoizeFunction"])((isGrabberClicked, isPreview) => {
    return Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["mergeStyles"])({
        alignItems: 'center',
        cursor: isPreview ? 'grabbing' : isGrabberClicked ? 'grab' : 'pointer',
        padding: 'unset',
        borderTopLeftRadius: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["grabberBorderRadius"],
        selectors: {
            ':focus': {
                outline: 'none'
            }
        }
    });
});
/**
 * The CSS class to apply to row grabber
 */
const grabberRowClassName = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["memoizeFunction"])((isGrabberClicked, rowHeight, isPreview) => {
    return Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["mergeStyles"])({
        width: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["rowGrabberWidth"],
        position: 'relative',
        height: rowHeight || 'auto',
        borderBottomLeftRadius: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["grabberBorderRadius"],
        // add bottom and left box-shadow to row grabber when selected to give multiple box-shadow effect.
        boxShadow: isGrabberClicked
            ? Object(_utils_getBoxShadow__WEBPACK_IMPORTED_MODULE_3__["getBoxShadow"])({
                color: grabberColorLabels.clickedBorder,
                type: _utils_getBoxShadow__WEBPACK_IMPORTED_MODULE_3__["boxShadowCellType"].rowGrabber,
                width: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["cellWrapperBorderWidth"]
            })
            : '',
        selectors: {
            [office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]]: {
                boxShadow: isGrabberClicked ? _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["rowHCBoxShadow"] : 'none'
            }
        }
    }, grabberCommonClassName(isGrabberClicked, isPreview));
});
/**
 * The CSS class to apply to column grabber
 */
const grabberColumnClassName = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["memoizeFunction"])((isGrabberClicked, columnWidth, isInitialCellWithSpecialHandling, isPreview, devicePixelRatio) => {
    const boxShadowWidth = devicePixelRatio < 1
        ? window.devicePixelRatio < 0.5
            ? window.devicePixelRatio < 0.34
                ? _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["cellWrapperBorderWidthZoomLessThanThird"]
                : _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["cellWrapperBorderWidthZoomLessThanHalf"]
            : _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["cellWrapperBorderWidthZoomLessThanOne"]
        : _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["cellWrapperBorderWidth"];
    return Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["mergeStyles"])({
        // unset position for initial Grabber Cell,
        // so that it is placed relatively to div outside Table.
        position: isInitialCellWithSpecialHandling ? 'unset' : 'relative',
        width: columnWidth,
        height: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["columnGrabberHeight"],
        borderTopRightRadius: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["grabberBorderRadius"],
        boxShadow: isGrabberClicked
            ? Object(_utils_getBoxShadow__WEBPACK_IMPORTED_MODULE_3__["getBoxShadow"])({
                color: grabberColorLabels.clickedBorder,
                type: _utils_getBoxShadow__WEBPACK_IMPORTED_MODULE_3__["boxShadowCellType"].columnGrabber,
                width: boxShadowWidth
            })
            : '',
        selectors: {
            [office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]]: {
                boxShadow: isGrabberClicked ? _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["columnGrabberHCBoxShadow"] : 'none'
            }
        }
    }, grabberCommonClassName(isGrabberClicked, isPreview));
});
/**
 * CSS style set for resize column callout.
 */
const resizeColumnCalloutStyleSet = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["mergeStyleSets"])({
    callout: {
        borderRadius: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].borderRadius,
        backgroundColor: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].backgroundColor,
        boxShadow: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].boxShadow,
        display: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["standardCSSDisplayTypes"].inlineBlock,
        height: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].height
    },
    text: {
        display: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["standardCSSDisplayTypes"].inline,
        margin: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].textMarginPadding,
        padding: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].textMarginPadding,
        fontFamily: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].textFontFamily,
        fontSize: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].replyAltIconSize,
        lineHeight: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].textLineHeight,
        color: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].textColor
    },
    textHolder: {
        display: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["standardCSSDisplayTypes"].flex,
        margin: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].textHolderMargin,
        paddingLeft: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].textHolderPaddingLeftRight,
        paddingRight: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].textHolderPaddingLeftRight,
        paddingTop: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].textHolderPaddingTopBottom,
        paddingBottom: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].textHolderPaddingTopBottom
    },
    fontIconLeft: {
        fontSize: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].replyAltIconSize,
        marginLeft: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].replyAltIconMarginLeftRight,
        marginRight: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].replyAltIconMarginLeftRight,
        marginTop: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].replyAltIconMarginTop,
        marginBottom: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].replyAltIconMarginBottom,
        verticalAlign: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].replyAltIconVerticalAlign
    },
    fontIconRight: {
        fontSize: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].replyAltIconSize,
        marginLeft: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].replyAltIconMarginLeftRight,
        marginRight: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].replyAltIconMarginLeftRight,
        marginTop: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].replyAltIconMarginTop,
        marginBottom: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].replyAltIconMarginBottom,
        verticalAlign: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].replyAltIconVerticalAlign,
        transform: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["resizeColumnCalloutStyleConstants"].replyAltIconTransform
    }
});


/***/ }),

/***/ "../tablero/lib/view/cells/RichTextCell.js":
/*!*************************************************!*\
  !*** ../tablero/lib/view/cells/RichTextCell.js ***!
  \*************************************************/
/*! exports provided: RichTextCell */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "RichTextCell", function() { return RichTextCell; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _table_Table_props__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../table/Table.props */ "../tablero/lib/view/table/Table.props.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/styling/7.13.5_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/styling/lib/index.js");
/* harmony import */ var _utilities_createComponent__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../../utilities/createComponent */ "../tablero/lib/utilities/createComponent.js");




const hostingDivClassName = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_2__["mergeStyles"])({
    textAlign: 'left'
});
// TODO: TASK 4016918: Refactor code so that we can reuse most of the component cell code
const RichTextCell = react__WEBPACK_IMPORTED_MODULE_0__["memo"](function RichTextCell(props) {
    const { rowId, columnId, columnType, onCellValueProposal, componentContext, hasLocalSelection, isReadOnlyMode } = props;
    const hostingRef = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](null);
    const loadRichTextComponent = () => {
        // tslint:disable-next-line: no-floating-promises
        Object(_utilities_createComponent__WEBPACK_IMPORTED_MODULE_3__["createComponentWithDataProvider"])(_table_Table_props__WEBPACK_IMPORTED_MODULE_1__["DataTypeMap"][columnType], componentContext.containerRuntime, {
            isTextBoxMode: true
        }).then((result) => {
            if (result) {
                onCellValueProposal(rowId, columnId, result);
            }
        });
    };
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        if (hasLocalSelection && !isReadOnlyMode) {
            // Do not let read only users create new components.
            loadRichTextComponent();
        }
    }, [hasLocalSelection, isReadOnlyMode]);
    return react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { ref: hostingRef, className: hostingDivClassName, tabIndex: 0 });
});


/***/ }),

/***/ "../tablero/lib/view/cells/TableCellWrapper.js":
/*!*****************************************************!*\
  !*** ../tablero/lib/view/cells/TableCellWrapper.js ***!
  \*****************************************************/
/*! exports provided: TableCellWrapper */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "TableCellWrapper", function() { return TableCellWrapper; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _StylingConstants__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../StylingConstants */ "../tablero/lib/view/StylingConstants.js");
/* harmony import */ var _StringTable__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../StringTable */ "../tablero/lib/view/StringTable.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/memoize.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/styling/7.13.5_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/styling/lib/index.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/KeyCodes.js");
/* harmony import */ var _presence_PresenceContainer__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../presence/PresenceContainer */ "../tablero/lib/view/presence/PresenceContainer.js");
/* harmony import */ var _columns_getColumnDragHandler__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../columns/getColumnDragHandler */ "../tablero/lib/view/columns/getColumnDragHandler.js");
/* harmony import */ var _ms_office_bohemia_ux__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! @ms/office-bohemia-ux */ "../office-bohemia-ux/lib/eventing/matchEventWithOrigin.js");
/* harmony import */ var _Constants__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ../Constants */ "../tablero/lib/view/Constants.js");
/* harmony import */ var _CellHelpers__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ./CellHelpers */ "../tablero/lib/view/cells/CellHelpers.js");
/* harmony import */ var _utils_getBoxShadow__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(/*! ../utils/getBoxShadow */ "../tablero/lib/view/utils/getBoxShadow.js");
/* harmony import */ var _table_TableAppFeatures__WEBPACK_IMPORTED_MODULE_12__ = __webpack_require__(/*! ../table/TableAppFeatures */ "../tablero/lib/view/table/TableAppFeatures.js");
/* harmony import */ var _zIndex__WEBPACK_IMPORTED_MODULE_13__ = __webpack_require__(/*! ../zIndex */ "../tablero/lib/view/zIndex.js");
/* harmony import */ var _utils_TableCellHighContrastUtility__WEBPACK_IMPORTED_MODULE_14__ = __webpack_require__(/*! ../utils/TableCellHighContrastUtility */ "../tablero/lib/view/utils/TableCellHighContrastUtility.js");
/* harmony import */ var _utils_isHighContrastModeOn__WEBPACK_IMPORTED_MODULE_15__ = __webpack_require__(/*! ../utils/isHighContrastModeOn */ "../tablero/lib/view/utils/isHighContrastModeOn.js");














const localSelectionColor = _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["colorLabels"].localSelection;
const localHoverColor = _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["colorLabels"].localHover;
// The HighContrastCellElement is lazily loaded when HC Mode is on, to improve performance.
const HighContrastCellElement = react__WEBPACK_IMPORTED_MODULE_0__["lazy"](() => __webpack_require__.e(/*! import() | HighContrastCellElement */ "HighContrastCellElement").then(__webpack_require__.bind(null, /*! ./HighContrastCellElement */ "../tablero/lib/view/cells/HighContrastCellElement.js")).then((module) => ({
    default: module.HighContrastCellElement
})));
/**
 * Used to create announcement for assistive technologies.
 */
var ColumnResizeAlert;
(function (ColumnResizeAlert) {
    ColumnResizeAlert[ColumnResizeAlert["None"] = 0] = "None";
    ColumnResizeAlert[ColumnResizeAlert["Increase"] = 1] = "Increase";
    ColumnResizeAlert[ColumnResizeAlert["Decrease"] = 2] = "Decrease";
})(ColumnResizeAlert || (ColumnResizeAlert = {}));
// The CSS class properties for <td>/<th>
const getCellWrapperClassName = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["memoizeFunction"])((selectionColor, cellBackground, isHeaderCell, styleOverride, showAttributionClickHighlight, showAttributionHoverHighlight) => Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_4__["mergeStyles"])({
    background: cellBackground,
    fontFamily: _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["fontFamily"],
    borderBottom: isHeaderCell ? `1px solid ${_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["colorLabels"].headerRowBottomBorderColor}` : 'none',
    // Show left and right borders for attribution highlight
    borderLeft: (showAttributionClickHighlight || showAttributionHoverHighlight) && !isHeaderCell
        ? `1px solid ${_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["colorLabels"].attributedCellBorderColor}`
        : `none`,
    borderRight: (showAttributionClickHighlight || showAttributionHoverHighlight) && !isHeaderCell
        ? `1px solid ${_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["colorLabels"].attributedCellBorderColor}`
        : `none`,
    position: 'relative',
    padding: 'unset',
    // Box shadow is used to provide a rounded outline to the cell when selected
    boxShadow: selectionColor === _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["standardColors"].transparent
        ? 'none'
        : `0 0 0px ${_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperBorderWidth"]}px ${selectionColor}, ${_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["selectedCellBoxShadow"]}`,
    // The selected cell's box shadow should appear above the neighboring cell's box shadow
    // z-index solves this issue by providing greater stack order.
    zIndex: selectionColor === _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["standardColors"].transparent ? 'unset' : _zIndex__WEBPACK_IMPORTED_MODULE_13__["selectedCellZIndex"],
    borderRadius: `${_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperBorderRadius"]}px`,
    height: '100%',
    verticalAlign: 'top',
    selectors: {
        ':focus': {
            // We are drawing our own border around the cell when it has focus, disable user agent default focus styles
            // TODO: TABLERO: Accessibility pass on the selection visuals
            outline: 'none'
        },
        ':hover': {
            // Box shadow is used to provide a rounded outline to the cell when hovered (with or without selection)
            boxShadow: selectionColor === _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["standardColors"].transparent
                ? `0 0 0px ${_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperBorderWidth"]}px ${localHoverColor}`
                : `0 0 0px ${_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperBorderWidth"]}px ${selectionColor}, ${_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["selectedCellBoxShadow"]}`,
            selectors: {
                [office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_4__["HighContrastSelector"]]: {
                    boxShadow: 'none',
                    outline: _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["highContrastOutline"],
                    MsHighContrastAdjust: 'none'
                }
            }
        },
        [office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_4__["HighContrastSelector"]]: {
            borderBottom: isHeaderCell ? '1px solid windowText' : 'none',
            background: 'window',
            // High contrast adjustment for attribution highlight
            border: !isHeaderCell && showAttributionClickHighlight
                ? '4px solid highlight'
                : !isHeaderCell && showAttributionHoverHighlight
                    ? '2px solid highlight'
                    : 'none',
            boxShadow: 'none',
            // give outline when there is any selection on cell.
            outline: selectionColor !== _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["standardColors"].transparent ? _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["highContrastOutline"] : 'unset',
            MsHighContrastAdjust: 'none'
        }
    }
}, 
// Will override box shadow if style override is passed in
styleOverride));
// The CSS class properties for cell contents (textarea, etc.)
const cellContentsClassName = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["memoizeFunction"])((isHeaderCell, width) => {
    const paddingTop = _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperPaddingTopAndBottom"];
    // Reduce bottom padding as there is a bottom border for header row
    const paddingBottom = isHeaderCell
        ? _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperPaddingTopAndBottom"] - _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["columnHeaderBottomElementHeight"]
        : _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperPaddingTopAndBottom"];
    // The padding is included in the height in border-box mode.
    const minHeight = _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["minComponentCellHeight"] + paddingBottom + paddingTop + 2 * _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperBorderWidth"];
    const minWidth = width ? (width < _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["minimumColumnWidth"] ? width : _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["minimumColumnWidth"]) : _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["minimumColumnWidth"];
    return Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_4__["mergeStyles"])({
        paddingLeft: _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperPaddingLeftAndRight"],
        paddingRight: _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperPaddingLeftAndRight"],
        paddingTop,
        paddingBottom,
        height: '100%',
        minHeight,
        boxSizing: 'border-box',
        minWidth
    });
});
// when the cell is focused, we want to stop the event from bubbling up (React focus bubbles unlike DOM focus events)
// so that no possible parent div focus handlers get triggered.
const onTableCellWrapperFocus = (ev) => {
    ev.stopPropagation();
};
/**
 * This is a component that wraps all table cells. It provides basic presence and focus behavior that is
 * consistent with the keyboard handlers registered at the table level.
 */
const TableCellWrapper = react__WEBPACK_IMPORTED_MODULE_0__["memo"](function TableCellWrapper(props) {
    const { rowId, columnId, columnIndex, hasLocalSelection, coauthorsWithSelection, onCellWrapperClick, children, width, isHeaderCell, onColumnResizeProposal, onUpdateCellSelection, onInvokeContextMenu, highlightForClickAttribution, highlightForHoverAttribution, cellStyleOverride, onResizeColumnLeaveCmd, isResizeColumnMode, onTableCellHover, isReadOnlyMode, tableCellApi, ariaMessage } = props;
    const columnResizeAlert = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](ColumnResizeAlert.None);
    let selectionColor = _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["standardColors"].transparent;
    let backgroundColor = _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["colorLabels"].cellBackground;
    const hasCoauthorsInCell = coauthorsWithSelection.length > 0;
    if (hasLocalSelection) {
        selectionColor = localSelectionColor;
    }
    else if (hasCoauthorsInCell) {
        selectionColor = coauthorsWithSelection[0].color;
    }
    if (highlightForClickAttribution) {
        backgroundColor = _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["attributedCellBackgroundForClick"];
    }
    else if (highlightForHoverAttribution) {
        backgroundColor = _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["attributedCellBackgroundForHover"];
    }
    const isCellSelected = hasLocalSelection || hasCoauthorsInCell;
    // Calculate net cell boundary override
    //
    // Say, for a given cell, the column Grabber is selected, and one of the insert icons in row grabber is hovered.
    // Cells in the former column will get override for all edges, and cells in the latter row will get override for one edge.
    // The intersecting cell will receive both these overrides. In this scenario,
    // it has to prioritize the all edges override over a single edge override.
    let styleOverride;
    if (Object(_table_TableAppFeatures__WEBPACK_IMPORTED_MODULE_12__["grabberFeatureEnabled"])()) {
        let boxShadow = '';
        if (cellStyleOverride) {
            if (cellStyleOverride === null || cellStyleOverride === void 0 ? void 0 : cellStyleOverride.allEdges) {
                // use default selectedCellBoxShadow for column's allEdges cellStyleOverride,
                // use only bottom box-shaow for row's allEdges cellStyleOverride.
                boxShadow = Object(_utils_getBoxShadow__WEBPACK_IMPORTED_MODULE_11__["getBoxShadow"])({
                    color: cellStyleOverride.allEdges.boundaryColor,
                    type: cellStyleOverride.allEdges.columnId
                        ? _utils_getBoxShadow__WEBPACK_IMPORTED_MODULE_11__["boxShadowCellType"].columnAllEdges
                        : _utils_getBoxShadow__WEBPACK_IMPORTED_MODULE_11__["boxShadowCellType"].rowBottomEdge,
                    width: _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperBorderWidth"]
                });
            }
            else if (cellStyleOverride.singleEdge) {
                if (cellStyleOverride.singleEdge.direction === 0 /* Left */) {
                    boxShadow = `-${_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperBorderWidth"]}px 0 0 0 ${cellStyleOverride.singleEdge.boundaryColor}, 0 0 0px ${_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperBorderWidth"]}px ${selectionColor}`;
                }
                if (cellStyleOverride.singleEdge.direction === 1 /* Right */) {
                    boxShadow = `${_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperBorderWidth"]}px 0 0 0 ${cellStyleOverride.singleEdge.boundaryColor}, 0 0 0px ${_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperBorderWidth"]}px ${selectionColor}`;
                }
                if (cellStyleOverride.singleEdge.direction === 2 /* Top */) {
                    boxShadow = `0 -${_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperBorderWidth"]}px 0 0 ${cellStyleOverride.singleEdge.boundaryColor}, 0 0 0px ${_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperBorderWidth"]}px ${selectionColor}`;
                }
                if (cellStyleOverride.singleEdge.direction === 3 /* Bottom */) {
                    boxShadow = `0 ${_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperBorderWidth"]}px 0 0 ${cellStyleOverride.singleEdge.boundaryColor}, 0 0 0px ${_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperBorderWidth"]}px ${selectionColor}`;
                }
            }
            styleOverride = {
                boxShadow,
                // Having cells close to each other and applying multiple box shadow on them may cause the shadow from an element to render on top of the previous element,
                // z-index solves this issue by providing greater stack order.
                zIndex: (cellStyleOverride === null || cellStyleOverride === void 0 ? void 0 : cellStyleOverride.allEdges) ? 1 : 'unset',
                selectors: {
                    ':hover, :hover:last-of-type': {
                        boxShadow,
                        selectors: {
                            [office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_4__["HighContrastSelector"]]: {
                                boxShadow: 'none',
                                // The Hovered cell's HC properties (like outline) should appear above cellStyleOverride properties of Column or Row Edges of same cell.
                                zIndex: _zIndex__WEBPACK_IMPORTED_MODULE_13__["selectedCellZIndex"]
                            }
                        }
                    },
                    ':last-of-type': {
                        // add a right boxShadow in addition to bottom box shadow for allEdges cellStyleOverride on last cell of a row.
                        boxShadow: (cellStyleOverride === null || cellStyleOverride === void 0 ? void 0 : cellStyleOverride.allEdges) && (cellStyleOverride === null || cellStyleOverride === void 0 ? void 0 : cellStyleOverride.allEdges.rowId)
                            ? Object(_utils_getBoxShadow__WEBPACK_IMPORTED_MODULE_11__["getBoxShadow"])({
                                color: cellStyleOverride.allEdges.boundaryColor,
                                type: _utils_getBoxShadow__WEBPACK_IMPORTED_MODULE_11__["boxShadowCellType"].rowLastEdge,
                                width: _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperBorderWidth"]
                            })
                            : boxShadow,
                        selectors: {
                            [office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_4__["HighContrastSelector"]]: {
                                // This selector selects the last cell element.
                                // For rows - No HC box shadow required for last cell element.
                                // For column - HC boxShadow is required on top edge of each cell (last column)
                                boxShadow: (cellStyleOverride === null || cellStyleOverride === void 0 ? void 0 : cellStyleOverride.allEdges) && (cellStyleOverride === null || cellStyleOverride === void 0 ? void 0 : cellStyleOverride.allEdges.columnId)
                                    ? Object(_utils_TableCellHighContrastUtility__WEBPACK_IMPORTED_MODULE_14__["getHCBoxShadowForGrabberSelection"])(isCellSelected, cellStyleOverride)
                                    : 'none'
                            }
                        }
                    },
                    [office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_4__["HighContrastSelector"]]: {
                        boxShadow: Object(_utils_TableCellHighContrastUtility__WEBPACK_IMPORTED_MODULE_14__["getHCBoxShadowForGrabberSelection"])(isCellSelected, cellStyleOverride),
                        // The Selected cell's HC properties (like outline) should appear above cellStyleOverride properties of Column or Row Edges of same cell.
                        zIndex: isCellSelected ? _zIndex__WEBPACK_IMPORTED_MODULE_13__["selectedCellZIndex"] : 'unset'
                    }
                },
                opacity: cellStyleOverride && cellStyleOverride.opacity ? cellStyleOverride.opacity.opacity : undefined
            };
        }
    }
    const onContextMenu = (event) => {
        if (onInvokeContextMenu) {
            // Don't show the browser's context menu.
            event.preventDefault();
            // Don't let any parent Tablero components handle the context menu.
            event.stopPropagation();
            // Callback to be called after every dismiss of contextMenu (dismiss can happen by different ways).
            // TODO: There can be other scenarios required to be handled, where cell restoration might be required like - scrolling, clicking outside the Table or on any other component, dismiss scenarios in mobile.
            const onDismissCallback = (event) => {
                // Restore the cell selection only in case when context menu is dismissed by escaping (from the context menu on the cell).
                if ((event === null || event === void 0 ? void 0 : event.nativeEvent) instanceof KeyboardEvent && (event === null || event === void 0 ? void 0 : event.nativeEvent.keyCode) === office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_5__["KeyCodes"].escape) {
                    onUpdateCellSelection();
                }
            };
            onInvokeContextMenu({ rowId, columnId, isHeaderCell, event, onDismissCallback });
        }
    };
    const getCoAuthorsAnnouncementString = (coauthors) => {
        const suffixString = 'also editing table cell';
        let coauthorsAnnouncementString = '';
        if (coauthors.length === 1) {
            coauthorsAnnouncementString = `${coauthors[0].name} is ${suffixString}`;
        }
        else if (coauthors.length === 2) {
            coauthorsAnnouncementString = `${coauthors[0].name} and ${coauthors[1].name} are ${suffixString}`;
        }
        else if (coauthors.length > 2) {
            coauthorsAnnouncementString = `${coauthors[0].name}, ${coauthors[1].name} and ${coauthors.length - 2} more user${coauthors.length - 2 > 1 ? 's' : ''} are ${suffixString}`;
        }
        return coauthorsAnnouncementString;
    };
    const colWidth = width ? width : 0;
    const columnWidthRef = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](colWidth);
    columnWidthRef.current = colWidth;
    const onResizeHandlerMouseDown = Object(_columns_getColumnDragHandler__WEBPACK_IMPORTED_MODULE_7__["getColumnDragHandler"])((deltaX) => {
        const newWidth = columnWidthRef.current + deltaX;
        // While decreasing column width, don't allow resizing if width becomes less than minimumColumnWidth
        if (deltaX < 0 && newWidth < _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["minimumColumnWidth"]) {
            return;
        }
        if (tableCellWrapperRef.current) {
            onColumnResizeProposal && onColumnResizeProposal(columnId, newWidth);
        }
    });
    const [isHighContrastModeStateOn, setHighContrastModeOn] = react__WEBPACK_IMPORTED_MODULE_0__["useState"](_utils_isHighContrastModeOn__WEBPACK_IMPORTED_MODULE_15__["isHighContrastModeOn"]);
    /**
     * adds event listener on highContrastMediaObserver and,
     * updates table's multiple background based upon isHighContrastModeOn flag.
     */
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        const updateTableHighContrastColor = () => {
            setHighContrastModeOn(_utils_isHighContrastModeOn__WEBPACK_IMPORTED_MODULE_15__["isHighContrastModeOn"]);
        };
        _utils_isHighContrastModeOn__WEBPACK_IMPORTED_MODULE_15__["highContrastMediaObserver"].addListener(updateTableHighContrastColor);
        return () => {
            _utils_isHighContrastModeOn__WEBPACK_IMPORTED_MODULE_15__["highContrastMediaObserver"].removeListener(updateTableHighContrastColor);
        };
    }, []);
    const tableCellWrapperRef = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](null);
    const columnResizeHandlerRef = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](null);
    const CellTag = isHeaderCell ? 'th' : 'td';
    const cellRole = isHeaderCell ? 'columnheader' : 'cell';
    // If we know the cell height add padding and border value to cover the spacing around the cell
    const cellResizeHandlerHeight = tableCellWrapperRef.current
        ? tableCellWrapperRef.current.clientHeight + _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperBorderWidth"] + 'px'
        : '100%';
    const getResizeHandlerClassName = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_4__["mergeStyles"])({
        width: _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["resizeHandlerWidth"],
        height: cellResizeHandlerHeight,
        position: 'absolute',
        right: -_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperBorderWidth"],
        top: 0,
        cursor: 'col-resize'
    });
    const onMouseDown = (ev) => {
        // TODO: This should be wrapped in a utility function that checks the event source and optionally takes a parameter of a callback to execute if the event is from hosted element
        const composedPath = ev.nativeEvent.composedPath ? ev.nativeEvent.composedPath() : Object(_ms_office_bohemia_ux__WEBPACK_IMPORTED_MODULE_8__["buildEventPath"])(ev.nativeEvent);
        if (Object(_ms_office_bohemia_ux__WEBPACK_IMPORTED_MODULE_8__["isEventFromComponentHostingElement"])(ev.currentTarget, composedPath)) {
            // If the click event is coming from with a hosted element in the cell make sure we restore the cell selection
            onUpdateCellSelection();
            return;
        }
        onCellWrapperClick();
    };
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        var _a;
        if (hasLocalSelection) {
            (_a = tableCellWrapperRef.current) === null || _a === void 0 ? void 0 : _a.scrollIntoView({
                behavior: 'smooth',
                block: 'nearest',
                inline: 'nearest'
            });
        }
    }, [hasLocalSelection]);
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        var _a;
        // Only run this effect when resize column from commanding menu feature is enabled.
        if (Object(_table_TableAppFeatures__WEBPACK_IMPORTED_MODULE_12__["resizeColumnFromCmdMenuFeatureEnabled"])()) {
            if (isResizeColumnMode) {
                (_a = columnResizeHandlerRef.current) === null || _a === void 0 ? void 0 : _a.focus();
            }
        }
    }, [isResizeColumnMode]);
    const onMouseEnter = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"](() => {
        // Don't invoke callback if grabber feature is disabled
        if (Object(_table_TableAppFeatures__WEBPACK_IMPORTED_MODULE_12__["grabberFeatureEnabled"])()) {
            onTableCellHover(columnId, rowId);
        }
    }, []);
    const onMouseLeave = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"](() => {
        // Don't invoke callback if grabber feature is disabled
        if (Object(_table_TableAppFeatures__WEBPACK_IMPORTED_MODULE_12__["grabberFeatureEnabled"])()) {
            onTableCellHover(/*columnId:*/ undefined, /*rowId:*/ undefined);
        }
    }, []);
    const onResizeColumnHandlerKeyDown = (ev) => {
        ev.stopPropagation();
        ev.preventDefault();
        if (columnResizeHandlerRef.current) {
            switch (ev.keyCode) {
                case office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_5__["KeyCodes"].left: {
                    const newWidth = colWidth - _Constants__WEBPACK_IMPORTED_MODULE_9__["columnWidthIncrement"];
                    // While decreasing column width, don't allow resizing if width becomes less than minimumColumnWidth
                    if (newWidth > _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["minimumColumnWidth"]) {
                        columnResizeAlert.current = ColumnResizeAlert.Decrease;
                        onColumnResizeProposal && onColumnResizeProposal(columnId, newWidth);
                    }
                    break;
                }
                case office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_5__["KeyCodes"].right: {
                    columnResizeAlert.current = ColumnResizeAlert.Increase;
                    onColumnResizeProposal && onColumnResizeProposal(columnId, colWidth + _Constants__WEBPACK_IMPORTED_MODULE_9__["columnWidthIncrement"]);
                    break;
                }
                default: {
                    // Any other keypress exits the user from resize mode - give focus and selection to the cell on which the resize column
                    // command was invoked. This inturn loses the column and grabber selection as well.
                    onResizeColumnHandlerBlur();
                    break;
                }
            }
        }
    };
    // This effect is used to create announcement for speech assistive technology.
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        if (columnResizeAlert.current !== ColumnResizeAlert.None && columnResizeHandlerRef.current) {
            ariaMessage(_StringTable__WEBPACK_IMPORTED_MODULE_2__["resizeAnnoucement"] +
                (columnIndex + 1) +
                (columnResizeAlert.current === ColumnResizeAlert.Increase
                    ? _StringTable__WEBPACK_IMPORTED_MODULE_2__["resizeAnnoucementIncrease"]
                    : _StringTable__WEBPACK_IMPORTED_MODULE_2__["resizeAnnoucementDecrease"]));
            // Update columnResizeAlert to avoid redundant announcements.
            columnResizeAlert.current = ColumnResizeAlert.None;
        }
    }, [columnResizeAlert.current, width]);
    const onResizeColumnHandlerBlur = () => {
        onResizeColumnLeaveCmd && onResizeColumnLeaveCmd();
    };
    const childDivRef = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](null);
    if (tableCellApi) {
        // API used by <Table> to clone preview of this component.
        tableCellApi[Object(_CellHelpers__WEBPACK_IMPORTED_MODULE_10__["getCellId"])(columnId, rowId)] = {
            getPreview: () => {
                let tableCellWrapperPreview;
                tableCellWrapperPreview = document.createElement(isHeaderCell ? 'th' : 'td');
                // Update style for this new element.
                tableCellWrapperPreview.classList.add(getCellWrapperClassName(selectionColor, backgroundColor, isHeaderCell, /*styleOverride*/ undefined));
                if (width) {
                    tableCellWrapperPreview.style.width = `${width}px`;
                }
                // Resize div and presence div are not needed for preview.
                if (childDivRef.current) {
                    tableCellWrapperPreview.appendChild(childDivRef.current.cloneNode(true));
                }
                return tableCellWrapperPreview;
            }
        };
    }
    return (
    // CellTag can be <td> or <th>
    react__WEBPACK_IMPORTED_MODULE_0__["createElement"](CellTag, { tabIndex: -1, onFocus: onTableCellWrapperFocus, onMouseEnter: onMouseEnter, onMouseLeave: onMouseLeave, key: columnId, className: getCellWrapperClassName(selectionColor, backgroundColor, isHeaderCell, styleOverride, highlightForClickAttribution, highlightForHoverAttribution), onMouseDown: onMouseDown, style: { width }, ref: tableCellWrapperRef, onContextMenu: onContextMenu, role: cellRole, "aria-label": getCoAuthorsAnnouncementString(coauthorsWithSelection) },
        isHighContrastModeStateOn && cellStyleOverride && (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](react__WEBPACK_IMPORTED_MODULE_0__["Suspense"], { fallback: null },
            react__WEBPACK_IMPORTED_MODULE_0__["createElement"](HighContrastCellElement, { cellStyleOverride: cellStyleOverride, isHeaderCell: isHeaderCell, tableCellWrapperRef: tableCellWrapperRef }))),
        !isReadOnlyMode && (react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { tabIndex: -1, className: getResizeHandlerClassName, ref: columnResizeHandlerRef, onMouseDown: onResizeHandlerMouseDown, onKeyDown: isResizeColumnMode ? onResizeColumnHandlerKeyDown : undefined, onBlur: isResizeColumnMode ? onResizeColumnHandlerBlur : undefined, style: {
                outline: 'none'
            }, "aria-hidden": true })),
        react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_presence_PresenceContainer__WEBPACK_IMPORTED_MODULE_6__["PresenceContainer"], { coauthors: coauthorsWithSelection, cellWidth: colWidth }),
        react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { className: cellContentsClassName(isHeaderCell, width), ref: childDivRef }, children)));
});


/***/ }),

/***/ "../tablero/lib/view/cells/TableDataCell.js":
/*!**************************************************!*\
  !*** ../tablero/lib/view/cells/TableDataCell.js ***!
  \**************************************************/
/*! exports provided: TableDataCell, ComponentCell */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "TableDataCell", function() { return TableDataCell; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ComponentCell", function() { return ComponentCell; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _presence_Presence__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../presence/Presence */ "../tablero/lib/view/presence/Presence.js");
/* harmony import */ var _FluidComponentCell__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./FluidComponentCell */ "../tablero/lib/view/cells/FluidComponentCell.js");
/* harmony import */ var _TableCellWrapper__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./TableCellWrapper */ "../tablero/lib/view/cells/TableCellWrapper.js");
/* harmony import */ var _dataModel_CellDataUtils__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../../dataModel/CellDataUtils */ "../tablero/lib/dataModel/CellDataUtils.js");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../../telemetry */ "../tablero/lib/telemetry/ErrorEvents.js");
/* harmony import */ var _fluidx_utilities__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! @fluidx/utilities */ "../utilities/lib/logger/LoggerHelpers.js");
/* harmony import */ var _dataModel_adapterComponents_TableroAdapterComponentFactory__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../../dataModel/adapterComponents/TableroAdapterComponentFactory */ "../tablero/lib/dataModel/adapterComponents/TableroAdapterComponentFactory.js");








const TableDataCell = react__WEBPACK_IMPORTED_MODULE_0__["memo"](function TableDataCell(props) {
    const { rowId, columnId, columnIndex, onCellSelectionProposal, coauthorsAtLocation, localPresenceLocation, width, onColumnResizeProposal, uiContextComponentUrl, onRowCompletionProposal, onInvokeContextMenu, onMoveSelectionFromCell, highlightForHoverAttribution, highlightForClickAttribution, cellStyleOverride, onTableCellHover, onResizeColumnLeaveCmd, isResizeColumnMode, logger, isReadOnlyMode, tableCellApi, ariaMessage } = props;
    const currentColumnLocation = { type: 'cell', rowId, columnId };
    const hasLocalSelection = localPresenceLocation !== undefined && Object(_presence_Presence__WEBPACK_IMPORTED_MODULE_1__["arePresenceLocationsEqual"])(currentColumnLocation, localPresenceLocation);
    const [cellWrapperFocused, setCellWrapperFocused] = react__WEBPACK_IMPORTED_MODULE_0__["useState"](false);
    /**
     * Invoked when the associated focus callback finished the execution.
     */
    const notifyCellWrapperFocusExecuted = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"](() => {
        setCellWrapperFocused(false);
    }, []);
    const onCellWrapperClick = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"](() => {
        setCellWrapperFocused(true);
        updateCellSelection();
    }, []);
    const updateCellSelection = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"](() => {
        onCellSelectionProposal(rowId, columnId);
    }, [rowId, columnId, onCellSelectionProposal]);
    return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_TableCellWrapper__WEBPACK_IMPORTED_MODULE_3__["TableCellWrapper"], { rowId: rowId, columnId: columnId, columnIndex: columnIndex, hasLocalSelection: hasLocalSelection, onCellWrapperClick: onCellWrapperClick, coauthorsWithSelection: coauthorsAtLocation || [], width: width, onColumnResizeProposal: onColumnResizeProposal, onUpdateCellSelection: updateCellSelection, onInvokeContextMenu: onInvokeContextMenu, highlightForClickAttribution: highlightForClickAttribution, highlightForHoverAttribution: highlightForHoverAttribution, cellStyleOverride: cellStyleOverride, onTableCellHover: onTableCellHover, isReadOnlyMode: isReadOnlyMode, onResizeColumnLeaveCmd: onResizeColumnLeaveCmd, isResizeColumnMode: isResizeColumnMode, tableCellApi: tableCellApi, ariaMessage: ariaMessage },
        react__WEBPACK_IMPORTED_MODULE_0__["createElement"](ComponentCell, Object.assign({}, props, { hasLocalSelection: hasLocalSelection, uiContextComponentUrl: uiContextComponentUrl, cellWrapperFocused: cellWrapperFocused, notifyCellWrapperFocusExecuted: notifyCellWrapperFocusExecuted, onUpdateCellSelection: updateCellSelection, onRowCompletionProposal: onRowCompletionProposal, onMoveSelectionFromCell: onMoveSelectionFromCell, logger: logger, tableCellApi: tableCellApi }))));
});
const ComponentCell = react__WEBPACK_IMPORTED_MODULE_0__["memo"](function renderComponentVisuals(props) {
    let cellData = props === null || props === void 0 ? void 0 : props.cellData;
    if (!cellData) {
        props.logger &&
            Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_6__["sendErrorEvent"])(props.logger, {
                eventName: _telemetry__WEBPACK_IMPORTED_MODULE_5__["ErrorEvent"].ComponentNestingError,
                errorCode: _telemetry__WEBPACK_IMPORTED_MODULE_5__["ErrorCode"].ComponentNestingError.CellDataUndefined
            });
        // Handling the case of props.cellData undefined by creating an adapter component.
        // TODO : Rethink N/A nested component scenarios handling.
        cellData = Object(_dataModel_adapterComponents_TableroAdapterComponentFactory__WEBPACK_IMPORTED_MODULE_7__["createAdapterComponent"])(props === null || props === void 0 ? void 0 : props.columnType);
    }
    const componentUrl = Object(_dataModel_CellDataUtils__WEBPACK_IMPORTED_MODULE_4__["tryGetComponentUrl"])(cellData);
    // If we don't have a componentUrl it means we use an adapter component.
    if (componentUrl === undefined) {
        const adapterVisualProvider = cellData.IComponentAdapterVisual;
        if (adapterVisualProvider === undefined) {
            // Logging the event when we failed to render the adapter component.
            props.logger &&
                Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_6__["sendErrorEvent"])(props.logger, {
                    eventName: _telemetry__WEBPACK_IMPORTED_MODULE_5__["ErrorEvent"].ComponentNestingError,
                    errorCode: _telemetry__WEBPACK_IMPORTED_MODULE_5__["ErrorCode"].ComponentNestingError.AdapterComponentUndefined
                });
            return react__WEBPACK_IMPORTED_MODULE_0__["createElement"](react__WEBPACK_IMPORTED_MODULE_0__["Fragment"], null);
        }
        return adapterVisualProvider.getAdapterComponentVisual(props);
    }
    return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_FluidComponentCell__WEBPACK_IMPORTED_MODULE_2__["FluidComponentCell"], { rowId: props.rowId, columnId: props.columnId, columnType: props.columnType, componentUrl: componentUrl, onCellValueProposal: props.onCellValueProposal, hasLocalSelection: props.hasLocalSelection, uiContextComponentUrl: props.uiContextComponentUrl, onUpdateCellSelection: props.onUpdateCellSelection, componentContext: props.componentContext, logger: props.logger, onMoveSelectionFromCell: props.onMoveSelectionFromCell, 
        // for the nested component, so that when nested component can't
        // handle caret position anymore, it lets table take control.
        onNestedComponentDataChangeCallback: props.onNestedComponentDataChangeCallback, isReadOnlyMode: props.isReadOnlyMode, isRowCompleted: props.isRowCompleted }));
});


/***/ }),

/***/ "../tablero/lib/view/cells/TextCell.js":
/*!*********************************************!*\
  !*** ../tablero/lib/view/cells/TextCell.js ***!
  \*********************************************/
/*! exports provided: TextCell */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "TextCell", function() { return TextCell; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _CellInput__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./CellInput */ "../tablero/lib/view/cells/CellInput.js");
/* harmony import */ var _dataModel_adapterComponents_TableroAdapterComponentFactory__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../../dataModel/adapterComponents/TableroAdapterComponentFactory */ "../tablero/lib/dataModel/adapterComponents/TableroAdapterComponentFactory.js");



const TextCell = react__WEBPACK_IMPORTED_MODULE_0__["memo"](function TextCell(props) {
    const { rowId, columnId, columnType, value, onCellValueProposal, hasLocalSelection, columnWidth, onUpdateCellSelection, sendInsertComponentRequestData, enableUndoRedo, componentContext, logger, onMoveSelectionFromCell, moveSelectionToCellDetails, isReadOnlyMode } = props;
    const onValueCommit = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((valueToCommit) => {
        const newCellData = Object(_dataModel_adapterComponents_TableroAdapterComponentFactory__WEBPACK_IMPORTED_MODULE_2__["createAdapterComponent"])(columnType, valueToCommit);
        onCellValueProposal(rowId, columnId, newCellData);
    }, [rowId, columnId, onCellValueProposal]);
    return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_CellInput__WEBPACK_IMPORTED_MODULE_1__["CellInput"], { rowId: rowId, columnId: columnId, columnType: columnType, committedValue: value, onValueCommit: onValueCommit, hasLocalSelection: hasLocalSelection, onCellValueProposal: onCellValueProposal, columnWidth: columnWidth, onUpdateCellSelection: onUpdateCellSelection, sendInsertComponentRequestData: sendInsertComponentRequestData, enableUndoRedo: enableUndoRedo, componentContext: componentContext, logger: logger, moveSelectionToCellDetails: hasLocalSelection ? moveSelectionToCellDetails : undefined, onMoveSelectionFromCell: onMoveSelectionFromCell, readOnly: isReadOnlyMode }));
});


/***/ }),

/***/ "../tablero/lib/view/columns/ColumnHeader.js":
/*!***************************************************!*\
  !*** ../tablero/lib/view/columns/ColumnHeader.js ***!
  \***************************************************/
/*! exports provided: ColumnHeader */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ColumnHeader", function() { return ColumnHeader; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _cells_TableCellWrapper__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../cells/TableCellWrapper */ "../tablero/lib/view/cells/TableCellWrapper.js");
/* harmony import */ var _cells_CellInput__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../cells/CellInput */ "../tablero/lib/view/cells/CellInput.js");
/* harmony import */ var _presence_Presence__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../presence/Presence */ "../tablero/lib/view/presence/Presence.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/styling/7.13.5_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/styling/lib/index.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/KeyCodes.js");
/* harmony import */ var _configuration_Presets__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../../configuration/Presets */ "../tablero/lib/configuration/Presets.js");
/* harmony import */ var _StylingConstants__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../StylingConstants */ "../tablero/lib/view/StylingConstants.js");
/* harmony import */ var _StringTable__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../StringTable */ "../tablero/lib/view/StringTable.js");








const columnHeaderWrapperClassName = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_4__["mergeStyles"])({
    display: 'flex',
    justifyContent: 'space-between',
    width: '100%',
    alignItems: 'center'
});
const ColumnHeader = react__WEBPACK_IMPORTED_MODULE_0__["memo"](function ColumnHeader(props) {
    const { title, id, columnIndex, width, onColumnSelected, onColumnTitleProposal, coauthorsInColumnHeader, localPresenceLocation, onColumnResizeProposal, onInvokeContextMenu, getPreset, sendInsertComponentRequestData, enableUndoRedo, componentContext, logger, onMoveSelectionFromCell, moveSelectionToCellDetails, cellStyleOverride, onResizeColumnLeaveCmd, isResizeColumnMode, onTableCellHover, isReadOnlyMode, tableCellApi, ariaMessage } = props;
    const currentColumnLocation = { type: 'column', columnId: id };
    const isColumnLocallySelected = localPresenceLocation !== undefined && Object(_presence_Presence__WEBPACK_IMPORTED_MODULE_3__["arePresenceLocationsEqual"])(currentColumnLocation, localPresenceLocation);
    const onColumnHeaderSelected = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"](() => onColumnSelected(id), [onColumnSelected, id]);
    const onColumnHeaderClick = () => {
        onColumnHeaderSelected();
    };
    const columnHeaderButtonRef = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](null);
    const onValueCommit = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((newColumnTitle) => onColumnTitleProposal(id, newColumnTitle), [
        id,
        onColumnTitleProposal
    ]);
    const onKeyDown = (ev) => {
        if (ev.altKey && ev.keyCode === office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_5__["KeyCodes"].down) {
            if (columnHeaderButtonRef.current) {
                columnHeaderButtonRef.current.focus();
            }
            ev.stopPropagation();
            return;
        }
    };
    const cellInputApi = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](null);
    return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_cells_TableCellWrapper__WEBPACK_IMPORTED_MODULE_1__["TableCellWrapper"], { columnId: id, columnIndex: columnIndex, onCellWrapperClick: onColumnHeaderClick, hasLocalSelection: isColumnLocallySelected, coauthorsWithSelection: coauthorsInColumnHeader || [], width: width, isHeaderCell: true, onColumnResizeProposal: onColumnResizeProposal, onUpdateCellSelection: onColumnHeaderSelected, onInvokeContextMenu: onInvokeContextMenu, cellStyleOverride: cellStyleOverride, onTableCellHover: onTableCellHover, isReadOnlyMode: isReadOnlyMode, onResizeColumnLeaveCmd: onResizeColumnLeaveCmd, 
        // Only the header cell having the matching column ID should receive true for resize table column mode.
        isResizeColumnMode: isResizeColumnMode, tableCellApi: tableCellApi, ariaMessage: ariaMessage },
        react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { className: columnHeaderWrapperClassName, onKeyDown: onKeyDown },
            react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_cells_CellInput__WEBPACK_IMPORTED_MODULE_2__["CellInput"], { columnId: id, committedValue: title, hasLocalSelection: isColumnLocallySelected, onValueCommit: onValueCommit, columnWidth: width, readOnly: getPreset(_configuration_Presets__WEBPACK_IMPORTED_MODULE_6__["PresetName"].DisableRenameColumns) || isReadOnlyMode, styleOverrides: {
                    fontWeight: _StylingConstants__WEBPACK_IMPORTED_MODULE_7__["columnHeaderFontWeight"],
                    fontFamily: _StylingConstants__WEBPACK_IMPORTED_MODULE_7__["columnHeaderFontFamily"],
                    fontSize: _StylingConstants__WEBPACK_IMPORTED_MODULE_7__["columnHeaderFontSize"],
                    lineHeight: _StylingConstants__WEBPACK_IMPORTED_MODULE_7__["columnHeaderLineHeight"],
                    flexGrow: 1
                }, sendInsertComponentRequestData: sendInsertComponentRequestData, enableUndoRedo: enableUndoRedo, componentContext: componentContext, logger: logger, cellInputApi: cellInputApi, onMoveSelectionFromCell: onMoveSelectionFromCell, moveSelectionToCellDetails: moveSelectionToCellDetails, additionalA11yLabel: _StringTable__WEBPACK_IMPORTED_MODULE_8__["headerCellAriaLabel"] }))));
});


/***/ }),

/***/ "../tablero/lib/view/columns/getColumnDragHandler.js":
/*!***********************************************************!*\
  !*** ../tablero/lib/view/columns/getColumnDragHandler.js ***!
  \***********************************************************/
/*! exports provided: getColumnDragHandler */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getColumnDragHandler", function() { return getColumnDragHandler; });
/**
 * Returns a mouse down handler to be used to handle column resize.
 */
function getColumnDragHandler(onDragMove) {
    // If there is a resize operation going on, this will be set with the last handled x position.
    let lastDragXPosition;
    const onMouseMove = (ev) => {
        if (lastDragXPosition) {
            onDragMove(ev.clientX - lastDragXPosition);
            lastDragXPosition = ev.clientX;
        }
    };
    const onMouseUp = (_) => {
        lastDragXPosition = undefined;
        document.removeEventListener('mousemove', onMouseMove);
        document.removeEventListener('mouseup', onMouseUp);
    };
    return (ev) => {
        lastDragXPosition = ev.clientX;
        document.addEventListener('mousemove', onMouseMove);
        document.addEventListener('mouseup', onMouseUp);
    };
}


/***/ }),

/***/ "../tablero/lib/view/header/TableHeader.js":
/*!*************************************************!*\
  !*** ../tablero/lib/view/header/TableHeader.js ***!
  \*************************************************/
/*! exports provided: TableHeader */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "TableHeader", function() { return TableHeader; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/styling/7.13.5_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/styling/lib/index.js");
/* harmony import */ var _TableTitle__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./TableTitle */ "../tablero/lib/view/header/TableTitle.js");
/* harmony import */ var _StylingConstants__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../StylingConstants */ "../tablero/lib/view/StylingConstants.js");




const tableHeaderWrapperClassName = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["mergeStyles"])({
    display: 'flex',
    justifyContent: 'space-between',
    height: `${_StylingConstants__WEBPACK_IMPORTED_MODULE_3__["tableHeaderHeight"]}px`,
    maxWidth: '100%'
});
const TableHeader = react__WEBPACK_IMPORTED_MODULE_0__["memo"](function TableHeader(props) {
    const { title, onTableTitleProposal, onTableTitleCommit, tableTitleId, tableTitleRef, onMoveToTableHeader, onMoveSelectionToHost, isReadOnlyMode } = props;
    return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](react__WEBPACK_IMPORTED_MODULE_0__["Fragment"], null,
        react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { className: tableHeaderWrapperClassName },
            react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_TableTitle__WEBPACK_IMPORTED_MODULE_2__["TableTitle"], { title: title, onTableTitleProposal: onTableTitleProposal, onTableTitleCommit: onTableTitleCommit, tableTitleId: tableTitleId, inputRef: tableTitleRef, onMoveToTableHeader: onMoveToTableHeader, onMoveSelectionToHost: onMoveSelectionToHost, isReadOnlyMode: isReadOnlyMode }))));
});


/***/ }),

/***/ "../tablero/lib/view/header/TableTitle.js":
/*!************************************************!*\
  !*** ../tablero/lib/view/header/TableTitle.js ***!
  \************************************************/
/*! exports provided: TableTitle */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "TableTitle", function() { return TableTitle; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/styling/7.13.5_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/styling/lib/index.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/KeyCodes.js");
/* harmony import */ var _StylingConstants__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../StylingConstants */ "../tablero/lib/view/StylingConstants.js");
/* harmony import */ var _zIndex__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../zIndex */ "../tablero/lib/view/zIndex.js");
/* harmony import */ var _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! @fluidx/office-fluid-types */ "../office-fluid-types/lib/ComponentContracts/ComponentSelection.js");





const wrapperClassName = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["mergeStyles"])({
    position: 'relative',
    display: 'flex',
    alignItems: 'center',
    flexGrow: 1,
    minWidth: _StylingConstants__WEBPACK_IMPORTED_MODULE_3__["tableTitleMinWidth"]
});
const inputClassName = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["mergeStyles"])({
    backgroundColor: _StylingConstants__WEBPACK_IMPORTED_MODULE_3__["colorLabels"].tableTitleBackground,
    width: '100%',
    border: 0,
    color: _StylingConstants__WEBPACK_IMPORTED_MODULE_3__["colorLabels"].tableTitleText,
    fontSize: _StylingConstants__WEBPACK_IMPORTED_MODULE_3__["titleFontSize"],
    fontFamily: _StylingConstants__WEBPACK_IMPORTED_MODULE_3__["titleFontFamily"],
    fontWeight: _StylingConstants__WEBPACK_IMPORTED_MODULE_3__["titleFontWeight"],
    lineHeight: _StylingConstants__WEBPACK_IMPORTED_MODULE_3__["titleLineHeight"],
    textOverflow: 'ellipsis',
    boxSizing: 'border-box',
    padding: 0,
    zIndex: _zIndex__WEBPACK_IMPORTED_MODULE_4__["tableTitleInputElementZIndex"],
    selectors: {
        ':focus': {
            outline: 'none'
        }
    }
});
/**
 * The component that contains the Table title and it's underline
 */
const TableTitle = react__WEBPACK_IMPORTED_MODULE_0__["memo"](function TableTitle(props) {
    const { title, onTableTitleProposal, onTableTitleCommit, tableTitleId, inputRef, onMoveToTableHeader, onMoveSelectionToHost, isReadOnlyMode } = props;
    const onChange = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((ev) => {
        onTableTitleProposal(ev.target.value);
    }, [onTableTitleProposal]);
    const onDragStart = (ev) => {
        ev.preventDefault();
    };
    const [inputHasFocus, setInputFocus] = react__WEBPACK_IMPORTED_MODULE_0__["useState"](false);
    const onFocus = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"](() => setInputFocus(true), []);
    const onBlur = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"](() => setInputFocus(false), []);
    const tryMoveSelectionToHost = () => {
        if (inputRef.current) {
            const currentSelectionStart = inputRef.current.selectionStart;
            const currentSelectionEnd = inputRef.current.selectionEnd;
            if (currentSelectionStart === currentSelectionEnd && currentSelectionStart === 0) {
                onMoveSelectionToHost({ mode: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_5__["SelectionMode"].ip, direction: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_5__["SelectionDirection"].left });
            }
        }
    };
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        // If the input just took focus, select all the text in it for easy title replacement.
        if (inputHasFocus && inputRef.current) {
            inputRef.current.select();
        }
    }, [inputHasFocus]);
    // Update the value of the input shown to reflect the new title value
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        if (inputRef.current) {
            inputRef.current.value = title;
            // For making it accessible for Read out reader
            inputRef.current.innerText = title;
        }
    }, [title]);
    const onKeyDown = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((ev) => {
        switch (ev.keyCode) {
            case office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_2__["KeyCodes"].enter:
                onTableTitleCommit();
                break;
            case office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_2__["KeyCodes"].down:
            case office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_2__["KeyCodes"].tab:
                onMoveToTableHeader();
                break;
            case office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_2__["KeyCodes"].left:
                tryMoveSelectionToHost();
                break;
            case office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_2__["KeyCodes"].up:
                ev.preventDefault();
                onMoveSelectionToHost({ direction: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_5__["SelectionDirection"].up });
        }
    }, [onTableTitleCommit]);
    return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { className: wrapperClassName },
        react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("input", { onChange: onChange, onKeyDown: onKeyDown, className: inputClassName, onFocus: onFocus, onBlur: onBlur, ref: inputRef, id: tableTitleId, readOnly: isReadOnlyMode, onDragStart: onDragStart })));
});


/***/ }),

/***/ "../tablero/lib/view/presence/Presence.js":
/*!************************************************!*\
  !*** ../tablero/lib/view/presence/Presence.js ***!
  \************************************************/
/*! exports provided: arePresenceLocationsEqual, addUserToMapInPlace, deleteUserFromMap, isPresenceInColumn, isPresenceInRow */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "arePresenceLocationsEqual", function() { return arePresenceLocationsEqual; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "addUserToMapInPlace", function() { return addUserToMapInPlace; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "deleteUserFromMap", function() { return deleteUserFromMap; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "isPresenceInColumn", function() { return isPresenceInColumn; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "isPresenceInRow", function() { return isPresenceInRow; });
function arePresenceLocationsEqual(presenceLocation1, presenceLocation2) {
    if (presenceLocation2 === undefined) {
        return false;
    }
    switch (presenceLocation1.type) {
        case 'cell': {
            return (presenceLocation2.type === 'cell' &&
                presenceLocation1.columnId === presenceLocation2.columnId &&
                presenceLocation1.rowId === presenceLocation2.rowId);
        }
        case 'column': {
            return presenceLocation2.type === 'column' && presenceLocation1.columnId === presenceLocation2.columnId;
        }
    }
}
/**
 * Adds a user to the coauthorsMap in-place.
 * This does not return a new map object.
 * Ensure to use it on a new map object for immutability.
 *
 * Current usage: FluidTable.tsx uses this after a deleteUserFromMap operation ensuring a new map object is passed in always
 */
function addUserToMapInPlace(currentMap, rowId, columnId, userInfo) {
    if (userInfo) {
        const rowValue = currentMap.get(rowId) || new Map();
        const cellValue = rowValue.get(columnId) || [];
        cellValue.push(userInfo);
        rowValue.set(columnId, cellValue);
        currentMap.set(rowId, rowValue);
    }
}
/**
 * Deletes a user from the coauthorsMap.
 * This returns a new map object.
 */
function deleteUserFromMap(coauthorsMap, userToDelete) {
    if (!userToDelete) {
        return coauthorsMap;
    }
    let newCoauthorsMap = new Map();
    coauthorsMap.forEach((columnMap, rowId) => {
        columnMap.forEach((users, columnId) => {
            users.forEach((user) => {
                // Skip adding the passed in user to the new Map
                if (user.userId !== userToDelete.userId) {
                    addUserToMapInPlace(newCoauthorsMap, rowId, columnId, user);
                }
            });
        });
    });
    return newCoauthorsMap;
}
/**
 * Returns true if the presence is in the current column, else false.
 */
function isPresenceInColumn(presenceLocation, columnId) {
    return (presenceLocation === null || presenceLocation === void 0 ? void 0 : presenceLocation.columnId) === columnId;
}
/**
 * Returns true if the presence is in the current row, else false.
 */
function isPresenceInRow(presenceLocation, rowId) {
    return (presenceLocation === null || presenceLocation === void 0 ? void 0 : presenceLocation.type) === 'cell' && presenceLocation.rowId === rowId;
}


/***/ }),

/***/ "../tablero/lib/view/presence/PresenceBubble.js":
/*!******************************************************!*\
  !*** ../tablero/lib/view/presence/PresenceBubble.js ***!
  \******************************************************/
/*! exports provided: PresenceBubble */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "PresenceBubble", function() { return PresenceBubble; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/styling/7.13.5_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/styling/lib/index.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/initials.js");
/* harmony import */ var _StylingConstants__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../StylingConstants */ "../tablero/lib/view/StylingConstants.js");
/* harmony import */ var _zIndex__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../zIndex */ "../tablero/lib/view/zIndex.js");
/* harmony import */ var _useAutoWidthAnimation__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ./useAutoWidthAnimation */ "../tablero/lib/view/presence/useAutoWidthAnimation.js");





// Class name for that makes up the bubble surrounding the user initials/name
const presenceBubbleWrapperClassName = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["mergeStyles"])({
    fontWeight: 400,
    display: 'block',
    overflowX: 'hidden',
    alignItems: 'center',
    position: 'absolute',
    height: _StylingConstants__WEBPACK_IMPORTED_MODULE_3__["presenceBubbleHeight"],
    borderRadius: 15,
    color: _StylingConstants__WEBPACK_IMPORTED_MODULE_3__["colorLabels"].presenceBubbleUserInitialsText,
    fontFamily: _StylingConstants__WEBPACK_IMPORTED_MODULE_3__["fontFamily"],
    zIndex: _zIndex__WEBPACK_IMPORTED_MODULE_4__["personaBubbleZIndex"],
    overflow: 'hidden',
    selectors: {
        [office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]]: {
            border: '1px solid windowText',
            boxSizing: 'border-box',
            MsHighContrastAdjust: 'auto'
        }
    }
});
const BubbleText = react__WEBPACK_IMPORTED_MODULE_0__["memo"](function CoauthorName(props) {
    const { coauthorName, isExpanded, overrideStyles } = props;
    const expandedFontSize = 12;
    const collapsedFontSize = 10;
    const spanClassName = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["mergeStyles"])({
        position: 'relative',
        overflow: 'hidden',
        display: 'block',
        whiteSpace: 'nowrap',
        margin: isExpanded ? '0 10px' : 0,
        fontSize: isExpanded ? expandedFontSize : collapsedFontSize
    });
    return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("span", { style: overrideStyles, className: spanClassName }, isExpanded ? coauthorName : Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_2__["getInitials"])(coauthorName, false /* isRtl */)));
});
/**
 * A control that signifies the identity of the user collaborating in the session
 */
const PresenceBubble = react__WEBPACK_IMPORTED_MODULE_0__["memo"](function PresenceBubble(props) {
    const { coauthorName, presenceColor, expandOnHover, cellWidth } = props;
    const [overrideStyles, setOverrideStyles] = react__WEBPACK_IMPORTED_MODULE_0__["useState"]({ textOverflow: 'unset' });
    const presenceBubbleRef = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](null);
    const { isExpanded, animationStyles, expandBubble, collapseBubble } = Object(_useAutoWidthAnimation__WEBPACK_IMPORTED_MODULE_5__["useAutoWidthAnimation"])({
        elementToAnimate: presenceBubbleRef,
        collapsedWidth: _StylingConstants__WEBPACK_IMPORTED_MODULE_3__["presenceBubbleCollapsedWidth"],
        maxWidth: cellWidth + _StylingConstants__WEBPACK_IMPORTED_MODULE_3__["cellWrapperBorderWidth"]
    });
    const onTransitionEnd = (_event) => {
        // We only want the text to hide text overflow with the ellipsis after it has finished the animation, and is expanded.
        // Otherwise the ellipsis would show for all presence bubbles until expanded regardless of length.
        if (isExpanded) {
            setOverrideStyles({ textOverflow: 'ellipsis' });
        }
    };
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        // This indicates the start of the the animation for the presence bubble collapsing. We do not want
        // to show the ellipsis during this time, or in a collapsed state.
        if (!isExpanded) {
            setOverrideStyles({ textOverflow: 'unset' });
        }
    }, [isExpanded]);
    return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](react__WEBPACK_IMPORTED_MODULE_0__["Fragment"], null,
        react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { className: presenceBubbleWrapperClassName, onMouseEnter: expandOnHover ? expandBubble : undefined, onMouseLeave: expandOnHover ? collapseBubble : undefined, onTransitionEnd: onTransitionEnd, ref: presenceBubbleRef, style: Object.assign({ display: 'flex', alignItems: 'center', justifyContent: 'center', backgroundColor: presenceColor }, animationStyles) },
            react__WEBPACK_IMPORTED_MODULE_0__["createElement"](BubbleText, { overrideStyles: overrideStyles, isExpanded: isExpanded, coauthorName: coauthorName }))));
});


/***/ }),

/***/ "../tablero/lib/view/presence/PresenceContainer.js":
/*!*********************************************************!*\
  !*** ../tablero/lib/view/presence/PresenceContainer.js ***!
  \*********************************************************/
/*! exports provided: PresenceContainer */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "PresenceContainer", function() { return PresenceContainer; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _PresenceBubble__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./PresenceBubble */ "../tablero/lib/view/presence/PresenceBubble.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/styling/7.13.5_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/styling/lib/index.js");
/* harmony import */ var _StylingConstants__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../StylingConstants */ "../tablero/lib/view/StylingConstants.js");




// Class name for a div that is solely about positioning the presence bubble relative to the cell
const presencePositionerClassName = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_2__["mergeStyles"])({
    position: 'absolute',
    height: _StylingConstants__WEBPACK_IMPORTED_MODULE_3__["presenceBubbleHeight"],
    width: _StylingConstants__WEBPACK_IMPORTED_MODULE_3__["presenceBubbleCollapsedWidth"]
});
// Offset between presence bubbles when they are stacked above one another
const presenceBubbleOffset = 4;
const PresenceContainer = react__WEBPACK_IMPORTED_MODULE_0__["memo"](function PresenceContainer(props) {
    const { coauthors, cellWidth } = props;
    const containerRef = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](null);
    return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { ref: containerRef }, coauthors.map((coauthor, index) => (react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { key: coauthor.userId, style: {
            top: -_StylingConstants__WEBPACK_IMPORTED_MODULE_3__["presenceBubbleHeight"] / 2,
            left: cellWidth - _StylingConstants__WEBPACK_IMPORTED_MODULE_3__["presenceBubbleCollapsedWidth"] - presenceBubbleOffset * index,
            zIndex: coauthors.length - index
        }, className: presencePositionerClassName }, _StylingConstants__WEBPACK_IMPORTED_MODULE_3__["presenceBubbleCollapsedWidth"] + presenceBubbleOffset * index <= cellWidth && (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_PresenceBubble__WEBPACK_IMPORTED_MODULE_1__["PresenceBubble"], { coauthorName: coauthor.name, presenceColor: coauthor.color, expandOnHover: index === 0, cellWidth: cellWidth })))))));
});


/***/ }),

/***/ "../tablero/lib/view/presence/useAutoWidthAnimation.js":
/*!*************************************************************!*\
  !*** ../tablero/lib/view/presence/useAutoWidthAnimation.js ***!
  \*************************************************************/
/*! exports provided: useAutoWidthAnimation */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "useAutoWidthAnimation", function() { return useAutoWidthAnimation; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);

function animationStateReducer(state, action) {
    switch (action.type) {
        // This should be called to trigger the animation sequence.
        case 'startExpand': {
            if (state.type !== 'collapsed') {
                return state;
            }
            // The first step in the animation sequence, where we signal that we want to measure the width of the div that will be animated.
            return { type: 'measureAutoWidth', styles: { width: 'auto' } };
        }
        // Called after we have gotten the expanded width and are ready to start animating.
        case 'prepareForAnimation': {
            return {
                type: 'prepareForAnimation',
                styles: action.collapsedStyles,
                expandedWidth: action.expandedWidth,
                left: action.left
            };
        }
        // Called after the styles have been reset to the collapsed state and we can start animating.
        case 'animateExpand': {
            if (state.type !== 'prepareForAnimation') {
                return state;
            }
            return {
                type: 'expanded',
                styles: {
                    width: state.expandedWidth,
                    left: state.left,
                    transition: 'width 300ms cubic-bezier(0.80, 0, 0.20, 1), left 300ms cubic-bezier(0.80, 0, 0.20, 1)'
                }
            };
        }
        // Called when we want to collapse the presence bubble back to the collapsed state.
        case 'collapse': {
            return {
                type: 'collapsed',
                styles: Object.assign(Object.assign({}, action.collapsedStyles), { transition: 'width 200ms cubic-bezier(0.10, 0.90, 0.20, 1), left 200ms cubic-bezier(0.10, 0.90, 0.20, 1)' })
            };
        }
    }
}
/**
 * A custom hook for animating elements that use a width of 'auto' to layout at a width that fits it's content.
 * This is to animate an element as it's width changes from a collapsed state with a fixed width to an expanded state with width auto.
 * This works around the fact that browser's don't have a way to animate width changes for elements that use width: auto.
 */
function useAutoWidthAnimation(payload) {
    const { elementToAnimate, collapsedWidth, maxWidth } = payload;
    // Styles to apply to the element when in the collapsed state
    const collapsedStyles = { width: collapsedWidth, left: 0 };
    // Summary of the animation code:
    // 1. Render the component in the collapsed state.
    // 2. User does something that would cause us to want to animate from collapsed to expanded.
    // 3. Render the element with width auto. Measure the width of the element in the expanded state.
    // 4. Render the element back in the collapsed state before #3 has painted. Set CSS transition on the width property.
    // 5. Set the width of the element to the value obtained in #3.
    //  Collapse is simple, just update transition and set the width back to the collapsed visuals
    const [animationState, dispatch] = react__WEBPACK_IMPORTED_MODULE_0__["useReducer"](animationStateReducer, {
        type: 'collapsed',
        styles: collapsedStyles
    });
    react__WEBPACK_IMPORTED_MODULE_0__["useLayoutEffect"](() => {
        // This hook gets called after layout, but before paint.
        // This enables us set width: 'auto' on the element to measure the width and then reset it back to the collapsed
        // state before the width: 'auto' render has been painted to the screen.
        if (animationState.type === 'measureAutoWidth' && elementToAnimate.current) {
            const elementWidth = elementToAnimate.current.getBoundingClientRect().width;
            const expandedWidth = elementWidth < maxWidth ? elementWidth : maxWidth;
            dispatch({
                type: 'prepareForAnimation',
                expandedWidth,
                left: -expandedWidth + collapsedWidth,
                collapsedStyles
            });
        }
    }, [animationState]);
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        // This hook gets called after the browser has painted.
        // This enables us to schedule the animation to start after we have reset the visuals to the collapsed state.
        if (animationState.type === 'prepareForAnimation') {
            dispatch({ type: 'animateExpand' });
        }
    }, [animationState]);
    return {
        isExpanded: animationState.type === 'measureAutoWidth' || animationState.type === 'expanded',
        animationStyles: animationState.styles,
        expandBubble: () => dispatch({ type: 'startExpand' }),
        collapseBubble: () => dispatch({ type: 'collapse', collapsedStyles })
    };
}


/***/ }),

/***/ "../tablero/lib/view/rows/AddRow.js":
/*!******************************************!*\
  !*** ../tablero/lib/view/rows/AddRow.js ***!
  \******************************************/
/*! exports provided: AddRow */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "AddRow", function() { return AddRow; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _table_Table_props__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../table/Table.props */ "../tablero/lib/view/table/Table.props.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/KeyCodes.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/styling/7.13.5_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/styling/lib/index.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/office-ui-fabric-react/7.121.6_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/office-ui-fabric-react/lib/components/Icon/FontIcon.js");
/* harmony import */ var _fluidx_icons__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! @fluidx/icons */ "../icons/lib/useIconNames/index.js");
/* harmony import */ var _StylingConstants__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../StylingConstants */ "../tablero/lib/view/StylingConstants.js");





const AddRow = react__WEBPACK_IMPORTED_MODULE_0__["memo"](function AddRow(props) {
    const { onInsertRow, rows, styleOverrides, label, newPresenceColumn, addRowBtnRef, setFocusAfterAddRowBtn } = props;
    let insertArgument;
    const insertNewRow = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"](() => {
        if (rows === null || rows === void 0 ? void 0 : rows.length) {
            insertArgument = {
                type: 'insertRelative',
                id: '',
                insertPosition: _table_Table_props__WEBPACK_IMPORTED_MODULE_1__["RelativeInsertPosition"].After,
                newPresenceLocation: newPresenceColumn
            };
            insertArgument.id = rows[rows.length - 1].id;
        }
        else {
            insertArgument = { type: 'insertStart' };
        }
        onInsertRow(insertArgument);
    }, [rows]);
    const onKeyPressedHandler = (event) => {
        if (event.keyCode === office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_2__["KeyCodes"].enter) {
            // Handling for adding a row as clicking on button
            insertNewRow();
        }
        else {
            setFocusAfterAddRowBtn(event.keyCode);
        }
        event.stopPropagation();
        return;
    };
    const parentClassName = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_3__["mergeStyles"])({
        height: '38px',
        marginTop: `-${_StylingConstants__WEBPACK_IMPORTED_MODULE_6__["tableBottomMargin"]}px`,
        fontSize: '14px',
        cursor: 'pointer',
        display: 'flex',
        alignItems: 'center',
        fontFamily: 'Segoe UI',
        selectors: {
            ':hover': {
                background: '#F3F2F1',
                outline: 'none'
            },
            ':focus': {
                outlineColor: '#000000'
            }
        }
    }, styleOverrides);
    const containerDivStyles = {
        display: 'flex',
        alignItems: 'baseline'
    };
    const addTaskBtnRole = 'button';
    return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { ref: addRowBtnRef, role: addTaskBtnRole, "aria-label": label, tabIndex: -1, onKeyDown: onKeyPressedHandler, onClick: insertNewRow, className: parentClassName },
        react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { style: containerDivStyles },
            react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { style: { marginRight: '14.75px', marginLeft: '14.75px', color: '#605E5C', fontSize: '12px' } },
                react__WEBPACK_IMPORTED_MODULE_0__["createElement"](office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_4__["FontIcon"], { style: { fontWeight: 'bold' }, iconName: Object(_fluidx_icons__WEBPACK_IMPORTED_MODULE_5__["useComponentsIcon"])('FFXCAdd') })),
            react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { style: { color: '#605E5C' } },
                " ",
                label,
                " "))));
});


/***/ }),

/***/ "../tablero/lib/view/rows/ColumnGrabberRow.js":
/*!****************************************************!*\
  !*** ../tablero/lib/view/rows/ColumnGrabberRow.js ***!
  \****************************************************/
/*! exports provided: ColumnGrabberRow */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ColumnGrabberRow", function() { return ColumnGrabberRow; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _cells_GrabberCell__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../cells/GrabberCell */ "../tablero/lib/view/cells/GrabberCell.js");
/* harmony import */ var _table_TableDefinitions__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../table/TableDefinitions */ "../tablero/lib/view/table/TableDefinitions.js");
/**
 * It maintains the collection of <td> column elements required by column Grabber Cell
 */



const ColumnGrabberRow = react__WEBPACK_IMPORTED_MODULE_0__["memo"](function ColumnGrabberRow(props) {
    const { onInsertColumnProposal, onDeleteColumnProposal, getPreset, onSortByColumnProposal, getLastKnownLocalPresence, onLocalPresenceValueProposal, onCellStyleOverride, hoveredTableCellColumnId, resizeColumnId, onResizeColumnLeaveCmd, tableWrapperElm, activateDragPreview, selectedColumnId, onSelect, isReadOnlyMode } = props;
    const onSelectColumn = (id) => {
        onSelect(id, _table_TableDefinitions__WEBPACK_IMPORTED_MODULE_2__["TableElementType"].Column);
    };
    const onMouseDown = (id) => {
        // User may drag grabber, which should result in
        // column re-order preview.
        activateDragPreview(_table_TableDefinitions__WEBPACK_IMPORTED_MODULE_2__["TableElementType"].Column, id);
    };
    return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("tr", null, props.columns.map((column, index) => {
        return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_cells_GrabberCell__WEBPACK_IMPORTED_MODULE_1__["GrabberCell"], { type: 1 /* Column */, key: column.id, id: column.id, isInitialCell: index === 0, isLastCell: index === props.columns.length - 1, columnWidth: column.width, onInsertProposal: onInsertColumnProposal, onDeleteProposal: onDeleteColumnProposal, getPreset: getPreset, sortDirection: column.sortDirection, onSortByColumnProposal: onSortByColumnProposal, getLastKnownLocalPresence: getLastKnownLocalPresence, onLocalPresenceValueProposal: onLocalPresenceValueProposal, onCellStyleOverride: onCellStyleOverride, hoveredTableCellId: hoveredTableCellColumnId, 
            // This will be true for the column for which resize column was invoked from the commanding menu.
            isColumnInResizeMode: resizeColumnId === column.id, onResizeColumnLeaveCmd: onResizeColumnLeaveCmd, tableWrapperElm: tableWrapperElm, onMouseDown: onMouseDown, isSelected: column.id === selectedColumnId, onSelect: onSelectColumn, isReadOnlyMode: isReadOnlyMode }));
    })));
});


/***/ }),

/***/ "../tablero/lib/view/rows/RowFiltering.js":
/*!************************************************!*\
  !*** ../tablero/lib/view/rows/RowFiltering.js ***!
  \************************************************/
/*! exports provided: filterRows, cellMatchesFilter, getFilterSuggestionsForColumn */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "filterRows", function() { return filterRows; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "cellMatchesFilter", function() { return cellMatchesFilter; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getFilterSuggestionsForColumn", function() { return getFilterSuggestionsForColumn; });
/* harmony import */ var _table_Table_props__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../table/Table.props */ "../tablero/lib/view/table/Table.props.js");
/* harmony import */ var _dataModel_CellDataUtils__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../../dataModel/CellDataUtils */ "../tablero/lib/dataModel/CellDataUtils.js");


/**
 * Filter out rows based on filter criterion specified per column
 */
const filterRows = async (rows, columns) => {
    const rowCellMismatchMarker = Symbol();
    // Transform rows by marking columns in row where cell fails to match filter
    const filterableRows = await Promise.all(rows.map(async (row) => {
        const filterColumns = await Promise.all(columns.map(async (column) => {
            const { id: columnId, filter } = column;
            if (filter !== undefined && !(await cellMatchesFilter(row.data[columnId], filter))) {
                return column;
            }
            return rowCellMismatchMarker;
        }));
        return { row, filterColumns };
    }));
    // Filter transformed rows
    const filteredRows = filterableRows.filter((filterableRow) => {
        const { filterColumns } = filterableRow;
        for (let column of filterColumns) {
            if (column !== rowCellMismatchMarker) {
                return false;
            }
        }
        return true;
    });
    // Remove column markings in filtered row and return
    return filteredRows.map((filterableRow) => filterableRow.row);
};
const cellMatchesFilter = async (cellData, filter) => {
    switch (filter.type) {
        case 'displayTextEquality': {
            const cellValue = await Object(_dataModel_CellDataUtils__WEBPACK_IMPORTED_MODULE_1__["getCellValue"])(cellData, _table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].Text);
            return cellValue === filter.filterValue;
        }
        case 'disjunction': {
            // If cell matches any of the filters within the disjunction then display the cell
            for (let filterItem of filter.operands) {
                if (await cellMatchesFilter(cellData, filterItem)) {
                    return true;
                }
            }
            return false;
        }
    }
};
const getFilterSuggestionsForColumn = async (cells) => {
    const textValueSet = new Set();
    await Promise.all(cells.map(async (cell) => {
        // Only components which are string producers can participate in table filtering.
        const cellValue = await Object(_dataModel_CellDataUtils__WEBPACK_IMPORTED_MODULE_1__["getCellValue"])(cell, _table_Table_props__WEBPACK_IMPORTED_MODULE_0__["DataType"].Text);
        if (cellValue !== undefined && typeof cellValue === 'string') {
            textValueSet.add(cellValue);
        }
    }));
    // The blank filter '' should be the last in the suggestions array.
    const blankSuggestion = textValueSet.delete('');
    const filterSuggestions = Array.from(textValueSet);
    if (blankSuggestion) {
        filterSuggestions.push('');
    }
    return filterSuggestions;
};


/***/ }),

/***/ "../tablero/lib/view/rows/TableRow.js":
/*!********************************************!*\
  !*** ../tablero/lib/view/rows/TableRow.js ***!
  \********************************************/
/*! exports provided: TableRow */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "TableRow", function() { return TableRow; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _presence_Presence__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../presence/Presence */ "../tablero/lib/view/presence/Presence.js");
/* harmony import */ var _cells_TableDataCell__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../cells/TableDataCell */ "../tablero/lib/view/cells/TableDataCell.js");
/* harmony import */ var _StylingConstants__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../StylingConstants */ "../tablero/lib/view/StylingConstants.js");
/* harmony import */ var _utils_Helper__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../utils/Helper */ "../tablero/lib/view/utils/Helper.js");
/* harmony import */ var _cells_CellHelpers__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../cells/CellHelpers */ "../tablero/lib/view/cells/CellHelpers.js");
/* harmony import */ var _utils_ResizeObserverWrapper__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../utils/ResizeObserverWrapper */ "../tablero/lib/view/utils/ResizeObserverWrapper.js");







const TableRow = react__WEBPACK_IMPORTED_MODULE_0__["memo"](function TableRow(props) {
    const { data, columns, onCellValueProposal, rowId, localPresenceLocation, onCellSelectionProposal, coauthorsInRow, onColumnResizeProposal, uiContextComponentUrl, onInvokeContextMenu, componentContext, logger, enableUndoRedo, sendInsertComponentRequestData, onMoveSelectionFromCell, moveSelectionToCellDetails, highlightColumnsForClickAttribution, highlightColumnsForHoverAttribution, cellStyleOverride, onTableCellHover, onNestedComponentDataChangeCallback, isReadOnlyMode, tableCellApi, updateRowHeights, onResizeColumnLeaveCmd, resizeColumnId, ariaMessage } = props;
    const ariaRowRole = 'row';
    const [isRowCompleted, setRowCompletion] = react__WEBPACK_IMPORTED_MODULE_0__["useState"](false);
    const onRowCompletionProposal = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((value) => {
        setRowCompletion(value);
    }, [setRowCompletion]);
    const isCellInAttributionList = (columnId, listToSearch) => {
        return listToSearch !== undefined && listToSearch.indexOf(columnId) >= 0;
    };
    const rowRef = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](null);
    /**
     * Attaching ResizeObserver for syncing the height of data table row with RowGrabbersTable row
     */
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        // tslint:disable-next-line: no-floating-promises
        Object(_utils_ResizeObserverWrapper__WEBPACK_IMPORTED_MODULE_6__["ensureResizeObserver"])().then(() => {
            if (!rowRef.current) {
                return;
            }
            const resizeObserver = new _utils_ResizeObserverWrapper__WEBPACK_IMPORTED_MODULE_6__["default"](() => {
                if (rowRef.current) {
                    updateRowHeights(rowId, rowRef.current.getBoundingClientRect().height);
                }
            });
            resizeObserver.observe(rowRef.current);
            return () => {
                // Remove all Elements from this observer
                resizeObserver.disconnect();
            };
        });
    }, [rowRef]);
    return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](react__WEBPACK_IMPORTED_MODULE_0__["Fragment"], null,
        react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("tr", { role: ariaRowRole, ref: rowRef }, columns.map((column, index) => (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_cells_TableDataCell__WEBPACK_IMPORTED_MODULE_2__["TableDataCell"], { key: column.id, columnIndex: index, rowId: rowId, columnId: column.id, columnType: column.type, cellData: data[column.id], onCellValueProposal: onCellValueProposal, 
            // Only the column that has local presence needs to know the location of local presence, others can receive undefined.
            localPresenceLocation: Object(_presence_Presence__WEBPACK_IMPORTED_MODULE_1__["isPresenceInColumn"])(localPresenceLocation, column.id) ? localPresenceLocation : undefined, onCellSelectionProposal: onCellSelectionProposal, coauthorsAtLocation: coauthorsInRow === null || coauthorsInRow === void 0 ? void 0 : coauthorsInRow.get(column.id), width: Object(_utils_Helper__WEBPACK_IMPORTED_MODULE_4__["getDimensionAsNumber"])(column.width, _StylingConstants__WEBPACK_IMPORTED_MODULE_3__["defaultColumnWidth"]), onColumnResizeProposal: onColumnResizeProposal, uiContextComponentUrl: uiContextComponentUrl, isRowCompleted: isRowCompleted, onRowCompletionProposal: onRowCompletionProposal, onInvokeContextMenu: onInvokeContextMenu, componentContext: componentContext, logger: logger, enableUndoRedo: enableUndoRedo, sendInsertComponentRequestData: sendInsertComponentRequestData, onMoveSelectionFromCell: onMoveSelectionFromCell, 
            // Only the column that has local presence needs to know the moveSelectionToCellDetails, others can receive undefined.
            moveSelectionToCellDetails: Object(_presence_Presence__WEBPACK_IMPORTED_MODULE_1__["isPresenceInColumn"])(localPresenceLocation, column.id) ? moveSelectionToCellDetails : undefined, highlightForClickAttribution: isCellInAttributionList(column.id, highlightColumnsForClickAttribution), highlightForHoverAttribution: isCellInAttributionList(column.id, highlightColumnsForHoverAttribution), cellStyleOverride: Object(_cells_CellHelpers__WEBPACK_IMPORTED_MODULE_5__["getCellStyleOverridesForCell"])(cellStyleOverride, column.id, rowId), onTableCellHover: onTableCellHover, tableCellApi: tableCellApi, onNestedComponentDataChangeCallback: onNestedComponentDataChangeCallback, isReadOnlyMode: isReadOnlyMode, onResizeColumnLeaveCmd: onResizeColumnLeaveCmd, 
            // Only the table cell having the matching column ID should receive true for resize table column mode
            isResizeColumnMode: resizeColumnId === column.id, ariaMessage: ariaMessage }))))));
});


/***/ }),

/***/ "../tablero/lib/view/table/ColumnGrabbersTable.js":
/*!********************************************************!*\
  !*** ../tablero/lib/view/table/ColumnGrabbersTable.js ***!
  \********************************************************/
/*! exports provided: ColumnGrabbersTable */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ColumnGrabbersTable", function() { return ColumnGrabbersTable; });
/* harmony import */ var _GrabberTableStyles__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./GrabberTableStyles */ "../tablero/lib/view/table/GrabberTableStyles.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_1___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_1__);
/* harmony import */ var _rows_ColumnGrabberRow__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../rows/ColumnGrabberRow */ "../tablero/lib/view/rows/ColumnGrabberRow.js");
/* harmony import */ var _TableStyle__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./TableStyle */ "../tablero/lib/view/table/TableStyle.js");




/**
 * ColumnGrabbersTable is a table above the data table, which has a column grabber for each data table column
 */
const ColumnGrabbersTable = react__WEBPACK_IMPORTED_MODULE_1__["memo"](function ColumnGrabbersTable(props) {
    const { columns, onInsertColumnProposal, onDeleteColumnProposal, getPreset, onSortByColumnProposal, getLastKnownLocalPresence, onLocalPresenceValueProposal, onCellStyleOverride, hoveredTableCellColumnId, tableWrapperElm, selectedColumnId, isReadOnlyMode, activateDragPreview, onRowColumnSelect, onResizeColumnLeaveCmd, resizeColumnId } = props;
    return (react__WEBPACK_IMPORTED_MODULE_1__["createElement"]("table", { className: `${Object(_GrabberTableStyles__WEBPACK_IMPORTED_MODULE_0__["columnGrabberTableClass"])()} ${_TableStyle__WEBPACK_IMPORTED_MODULE_3__["tableBorderSpacingClassName"]}` },
        react__WEBPACK_IMPORTED_MODULE_1__["createElement"]("tbody", null,
            react__WEBPACK_IMPORTED_MODULE_1__["createElement"](_rows_ColumnGrabberRow__WEBPACK_IMPORTED_MODULE_2__["ColumnGrabberRow"], { columns: columns, onInsertColumnProposal: onInsertColumnProposal, onDeleteColumnProposal: onDeleteColumnProposal, onSortByColumnProposal: onSortByColumnProposal, getPreset: getPreset, getLastKnownLocalPresence: getLastKnownLocalPresence, onLocalPresenceValueProposal: onLocalPresenceValueProposal, onCellStyleOverride: onCellStyleOverride, hoveredTableCellColumnId: hoveredTableCellColumnId, tableWrapperElm: tableWrapperElm, selectedColumnId: selectedColumnId, onSelect: onRowColumnSelect, isReadOnlyMode: isReadOnlyMode, activateDragPreview: activateDragPreview, onResizeColumnLeaveCmd: onResizeColumnLeaveCmd, resizeColumnId: resizeColumnId }))));
});


/***/ }),

/***/ "../tablero/lib/view/table/FluidTable.js":
/*!***********************************************!*\
  !*** ../tablero/lib/view/table/FluidTable.js ***!
  \***********************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @fluidx/office-fluid-types */ "../office-fluid-types/lib/ContainerServices/PresenceManager.js");
/* harmony import */ var _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../../view/table/Table.props */ "../tablero/lib/view/table/Table.props.js");
/* harmony import */ var _view_table_Table__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../../view/table/Table */ "../tablero/lib/view/table/Table.js");
/* harmony import */ var _dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../../dataModel/TableroAdapter */ "../tablero/lib/dataModel/TableroAdapter.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/array.js");
/* harmony import */ var _dataModel_CellDataUtils__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../../dataModel/CellDataUtils */ "../tablero/lib/dataModel/CellDataUtils.js");
/* harmony import */ var _presence_Presence__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../presence/Presence */ "../tablero/lib/view/presence/Presence.js");
/* harmony import */ var _configuration_Presets__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../../configuration/Presets */ "../tablero/lib/configuration/Presets.js");
/* harmony import */ var _configuration_ConfigurationUtils__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ../../configuration/ConfigurationUtils */ "../tablero/lib/configuration/ConfigurationUtils.js");
/* harmony import */ var _commanding_getCommands__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ../../commanding/getCommands */ "../tablero/lib/commanding/getCommands.js");
/* harmony import */ var _dataModel_TableroStore__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(/*! ../../dataModel/TableroStore */ "../tablero/lib/dataModel/TableroStore.js");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_12__ = __webpack_require__(/*! ../../telemetry */ "../tablero/lib/telemetry/ErrorEvents.js");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_13__ = __webpack_require__(/*! ../../telemetry */ "../tablero/lib/telemetry/UserAction.js");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_14__ = __webpack_require__(/*! ../../telemetry */ "../tablero/lib/telemetry/Activities.js");
/* harmony import */ var _fluidx_utilities__WEBPACK_IMPORTED_MODULE_15__ = __webpack_require__(/*! @fluidx/utilities */ "../utilities/lib/logger/LoggerHelpers.js");
/* harmony import */ var _fluidx_utilities__WEBPACK_IMPORTED_MODULE_16__ = __webpack_require__(/*! @fluidx/utilities */ "../utilities/lib/fluidExtension/MapRevertible.js");
/* harmony import */ var _fluidx_utilities__WEBPACK_IMPORTED_MODULE_17__ = __webpack_require__(/*! @fluidx/utilities */ "../utilities/lib/fluidExtension/SequenceRevertible.js");
/* harmony import */ var _fluidx_utilities__WEBPACK_IMPORTED_MODULE_18__ = __webpack_require__(/*! @fluidx/utilities */ "../utilities/lib/logger/ActivityTracker.js");
/* harmony import */ var _fluidx_utilities__WEBPACK_IMPORTED_MODULE_19__ = __webpack_require__(/*! @fluidx/utilities */ "../utilities/lib/presenceColorProvider/resolvePresenceColor.js");
/* harmony import */ var _utilities_stylingUtils__WEBPACK_IMPORTED_MODULE_20__ = __webpack_require__(/*! ../../utilities/stylingUtils */ "../tablero/lib/utilities/stylingUtils.js");
/* harmony import */ var _Constants__WEBPACK_IMPORTED_MODULE_21__ = __webpack_require__(/*! ../Constants */ "../tablero/lib/view/Constants.js");
/* harmony import */ var _StylingConstants__WEBPACK_IMPORTED_MODULE_22__ = __webpack_require__(/*! ../StylingConstants */ "../tablero/lib/view/StylingConstants.js");
/* harmony import */ var _configuration_Configuration__WEBPACK_IMPORTED_MODULE_23__ = __webpack_require__(/*! ../../configuration/Configuration */ "../tablero/lib/configuration/Configuration.js");
/* harmony import */ var _StringTable__WEBPACK_IMPORTED_MODULE_24__ = __webpack_require__(/*! ../StringTable */ "../tablero/lib/view/StringTable.js");
/* harmony import */ var _utilities_attributionUtils__WEBPACK_IMPORTED_MODULE_25__ = __webpack_require__(/*! ../../utilities/attributionUtils */ "../tablero/lib/utilities/attributionUtils.js");
/* harmony import */ var _dataModel_TableroDocumentUtil__WEBPACK_IMPORTED_MODULE_26__ = __webpack_require__(/*! ../../dataModel/TableroDocumentUtil */ "../tablero/lib/dataModel/TableroDocumentUtil.js");
/* harmony import */ var _TableAppFeatures__WEBPACK_IMPORTED_MODULE_27__ = __webpack_require__(/*! ./TableAppFeatures */ "../tablero/lib/view/table/TableAppFeatures.js");
/* harmony import */ var _fluidx_icons__WEBPACK_IMPORTED_MODULE_28__ = __webpack_require__(/*! @fluidx/icons */ "../icons/lib/iconSets/components.js");
























// Event listener names
const valueChanged = 'valueChanged';
const sequenceDelta = 'sequenceDelta';
const removeMember = 'removeMember';
const beforeunload = 'beforeunload';
const readonly = 'readonly';
const signal = 'signal';
const connected = 'connected';
// Signal types
const presenceSignal = 'presence';
const requestPresenceSignal = 'requestPresence';
const componentFriendlyName = 'componentFriendlyName';
window.FluidTable = class FluidTable extends react__WEBPACK_IMPORTED_MODULE_0__["PureComponent"] {
    constructor(props) {
        super(props);
        // Factories for the commanding surface commands. Two dimensional so we know which command
        // groups the created commands belong into.
        this.commandFactories = [];
        /**
         * Listener function to be called when theme properties are changed
         */
        this.themeChangeListener = (previousThemeCssProperties) => {
            Object(_utilities_stylingUtils__WEBPACK_IMPORTED_MODULE_20__["clearAndUpdateThemeVariables"])(previousThemeCssProperties, this.wrapperRef.current, this.stylingService);
        };
        /**
         * Helper function to find first missing elements in row/column.
         * We use this function only in case of deletion thus oldArr is larger than newArr.
         */
        this.firstMissingElem = (oldArr, newArr) => {
            return oldArr.findIndex((oldItem) => !newArr.some((newItem) => oldItem.id === newItem.id));
        };
        this.registerEventListeners = () => {
            const { componentContext, enableUndoRedo, logger, presenceManager, componentRuntime } = this.props.model;
            const { tableroViewStore, tableroDocumentStore } = this.store;
            // TODO: TRANSACTIONAL OPS: Do we want to only listen on explicit rerender ops instead of all rootmap changes?
            tableroViewStore.rootMap.on(valueChanged, this.viewStoreRootMapValueChangedListener);
            tableroDocumentStore.rootMap.on(valueChanged, this.documentStoreRootMapValueChangedListener); // TODO: Temporary for transactional ops from table document
            if (presenceManager === undefined) {
                componentContext.getAudience().on(removeMember, this.removeMemberForPresence);
            }
            window.addEventListener(beforeunload, this.onWindowBeforeUnload);
            if (enableUndoRedo) {
                // Listen to sequence delta event for undo/redo purpose.
                // TODO:COMPONENTS: When we have multiple copies of same Tablero, do not register the event handler multiple times.
                if (tableroDocumentStore.table.sequenceDeltaListened === undefined) {
                    tableroDocumentStore.table.sequenceDeltaListened = 0;
                }
                if (tableroDocumentStore.table.sequenceDeltaListened === 0) {
                    // TODO:UNDO: This need to be changed if we turn on table document.
                    tableroDocumentStore.table.on(sequenceDelta, this.onReceivedSequenceDeltaOp);
                }
                tableroDocumentStore.table.sequenceDeltaListened += 1;
            }
            // Update state for read only mode in case delta manager fires readonly event.
            componentRuntime.deltaManager.on(readonly, this.onReadOnlyEvent);
            try {
                // Listen to signals
                componentRuntime.on(signal, (message, local) => {
                    // Ignore own signals.
                    if (local) {
                        return;
                    }
                    switch (message.type) {
                        case presenceSignal:
                            this.onPresenceSignalReceived(message);
                            break;
                        case requestPresenceSignal:
                            this.onRequestPresenceSignalReceived(message);
                            break;
                    }
                });
                // Ask remote clients to send their presence, so we know where they are
                componentRuntime.on(connected, () => {
                    if (presenceManager !== undefined) {
                        // For compatibility, we still need to ask the clients that don't have containerPresenceManager to tell us their presence location.
                        // TODO: Remove following code when max unsupported version have been bumped to a version that contains containerPresenceManager implementation.
                        componentRuntime.submitSignal(requestPresenceSignal, {
                            forLegacyClients: true // set a flag to indicate this signal is sent from a new client just for compatibility.
                        });
                    }
                    else {
                        componentRuntime.submitSignal(requestPresenceSignal, undefined);
                    }
                });
            }
            catch (error) {
                logger &&
                    Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_15__["sendErrorEvent"])(logger, {
                        eventName: _telemetry__WEBPACK_IMPORTED_MODULE_12__["ErrorEvent"].PresenceSignalSubmissionError,
                        errorCode: _telemetry__WEBPACK_IMPORTED_MODULE_12__["ErrorEvent"].PresenceSignalSubmissionError
                    }, error);
            }
        };
        this.onWindowBeforeUnload = (_event) => {
            // Clear the current client's presence.
            this.broadcastLocalPresence(undefined);
        };
        this.removeMemberForPresence = (clientId) => {
            const { componentContext } = this.props.model;
            if (clientId === componentContext.clientId) {
                return;
            }
            const coauthorsMap = Object(_presence_Presence__WEBPACK_IMPORTED_MODULE_7__["deleteUserFromMap"])(this.state.coauthorsMap, this.getBasicUserInfo(clientId));
            this.setState({ coauthorsMap });
        };
        this.onRootMapValueChanged = async (changed, isLocal, target) => {
            const { enableUndoRedo, componentContext } = this.props.model;
            const tableroStore = this.store;
            if (enableUndoRedo && isLocal) {
                const revertible = new _fluidx_utilities__WEBPACK_IMPORTED_MODULE_16__["MapRevertible"](changed, target);
                if (!tableroStore.revertibleArray) {
                    tableroStore.revertibleArray = [];
                }
                const notExists = tableroStore.revertibleArray.every((item) => {
                    if (item instanceof _fluidx_utilities__WEBPACK_IMPORTED_MODULE_16__["MapRevertible"]) {
                        const mapRevertible = item;
                        return !mapRevertible.equals(revertible);
                    }
                    return true;
                });
                if (notExists) {
                    tableroStore.revertibleArray.push(revertible);
                }
            }
            const viewTableTitle = tableroStore.tableroViewStore.rootMap.get(componentFriendlyName);
            if (viewTableTitle !== tableroStore.tableroDocumentStore.tableTitle) {
                tableroStore.tableroDocumentStore.tableTitle = viewTableTitle;
                if (this.tableroDocument === undefined) {
                    this.tableroDocument = await this.getTableroDocument(tableroStore, componentContext);
                }
                this.tableroDocument.notifyTableUpdate();
            }
            // tslint:disable-next-line: no-floating-promises
            this.onDocumentStoreRootMapValueChanged();
        };
        // Update presence to first cell of previous row or first cell in case of no previous row exists.
        this.updatePresencePostRowDeletion = (indexOfDeletedRow) => {
            const { rows, columns } = this.state;
            let presence;
            if (rows.length > 1) {
                // indexOfPrevRow - row index of previous row or next row if next (0th) row is deleted.
                const indexOfPrevRow = indexOfDeletedRow === 0 ? 1 : indexOfDeletedRow - 1;
                presence = { type: 'cell', columnId: columns[0].id, rowId: rows[indexOfPrevRow].id };
            }
            else {
                // Select first column (Header Row) when no data rows present.
                presence = { type: 'column', columnId: columns[0].id };
            }
            // tslint:disable-next-line: no-floating-promises
            this.onLocalPresenceValueProposal(presence);
        };
        // Update presence to previous column (cell) or first column in case of deletion of first column itself.
        this.updatePresencePostColumnDeletion = (indexOfDeletedColumn, columns) => {
            if (columns.length > 1) {
                // indexOfPrevColumn - column index of previous row or next column incase first row was deleted.
                const indexOfPrevColumn = indexOfDeletedColumn === 0 ? 1 : indexOfDeletedColumn - 1;
                // tslint:disable-next-line: no-floating-promises
                this.onLocalPresenceValueProposal({ type: 'column', columnId: columns[indexOfPrevColumn].id });
            }
            else {
                // TODO: Check how can we transfer presence to header of Task/Table.
            }
        };
        this.viewStoreRootMapValueChangedListener = (changed, isLocal) => {
            // tslint:disable-next-line: no-floating-promises
            this.onRootMapValueChanged(changed, isLocal, this.store.tableroViewStore.rootMap);
        };
        this.documentStoreRootMapValueChangedListener = () => {
            // tslint:disable-next-line: no-floating-promises
            this.onDocumentStoreRootMapValueChanged();
        };
        this.onReceivedSequenceDeltaOp = (deltaEvent, target) => {
            if (deltaEvent.isLocal) {
                const revertible = new _fluidx_utilities__WEBPACK_IMPORTED_MODULE_17__["SequenceRevertible"](deltaEvent, target);
                const tableroStore = this.store;
                if (!tableroStore.revertibleArray) {
                    tableroStore.revertibleArray = [];
                }
                tableroStore.revertibleArray.push(revertible);
            }
        };
        this.onReadOnlyEvent = (isReadonly) => {
            this.setState({ isReadOnlyMode: isReadonly });
        };
        this.onCellValueProposal = (rowId, columnId, newCellData) => {
            this.store.postTransaction(async () => {
                await Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["setCellValue"])(this.store.tableroDocumentStore, rowId, columnId, newCellData, this.props.model.logger);
            }, _telemetry__WEBPACK_IMPORTED_MODULE_13__["UserAction"].CommitTableCellValue, _dataModel_TableroStore__WEBPACK_IMPORTED_MODULE_11__["EditMode"].Edit, true /* setNoop */);
        };
        this.onColumnResizeProposal = (columnId, newWidth) => {
            this.store.postTransaction(() => {
                Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["updateColumnViewProperties"])(this.store.tableroViewStore, columnId, { width: newWidth });
            }, this.state.resizeColumnProperties.columnId
                ? _telemetry__WEBPACK_IMPORTED_MODULE_13__["UserAction"].ResizeTableColumnFromCommandingMenu
                : _telemetry__WEBPACK_IMPORTED_MODULE_13__["UserAction"].ResizeTableColumn, _dataModel_TableroStore__WEBPACK_IMPORTED_MODULE_11__["EditMode"].Edit);
        };
        this.onResizeColumnCmd = (rowId, columnId) => {
            this.setState({ resizeColumnProperties: { rowId, columnId } });
        };
        this.onSortByColumnProposal = (sortByColumnId, sortDirection, requestOrigin) => {
            this.store.postTransaction(() => {
                Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["setSortByColumn"])(this.store.tableroViewStore, {
                    sortByColumnId,
                    sortDirection,
                    applySort: true
                });
            }, _telemetry__WEBPACK_IMPORTED_MODULE_13__["UserAction"].SortTableByColumn, _dataModel_TableroStore__WEBPACK_IMPORTED_MODULE_11__["EditMode"].Edit, undefined /* setNoop */, requestOrigin);
        };
        this.onFilterColumnProposal = (columnId, newFilterState) => {
            this.store.postTransaction(() => {
                const newFilteredColumnIds = Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getHiddenColumnIds"])(this.store.tableroViewStore).filter((id) => id !== columnId);
                if (newFilterState) {
                    newFilteredColumnIds.push(columnId);
                }
                Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["setHiddenColumnIds"])(this.store.tableroViewStore, newFilteredColumnIds);
                this.setFocusOnNextFocussableColumn(columnId);
            }, _telemetry__WEBPACK_IMPORTED_MODULE_13__["UserAction"].FilterTableByColumn, _dataModel_TableroStore__WEBPACK_IMPORTED_MODULE_11__["EditMode"].Edit);
        };
        this.onShowAllColumnsProposal = () => {
            this.store.postTransaction(() => {
                Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["setHiddenColumnIds"])(this.store.tableroViewStore, []);
            }, _telemetry__WEBPACK_IMPORTED_MODULE_13__["UserAction"].UnhideTableColumns, _dataModel_TableroStore__WEBPACK_IMPORTED_MODULE_11__["EditMode"].Edit);
        };
        this.onInsertRowProposal = (arg, requestOrigin) => {
            this.store.postTransaction(async () => {
                const { tableroDocumentStore, tableroViewStore } = this.store;
                const orderedRowIds = Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getRowViewProperties"])(tableroViewStore);
                let indexToInsertDM = 0;
                let indexToInsertVM = 0;
                switch (arg.type) {
                    case 'insertRelative': {
                        const indexOfRow = tableroDocumentStore.rowIdToIndex[arg.id];
                        if (indexOfRow !== undefined) {
                            indexToInsertDM = arg.insertPosition === _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__["RelativeInsertPosition"].After ? indexOfRow + 1 : indexOfRow;
                        }
                        indexToInsertVM = Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getOrderedRowIndexFromId"])(tableroViewStore, arg.id);
                        if (arg.insertPosition === _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__["RelativeInsertPosition"].After) {
                            indexToInsertVM = indexToInsertVM + 1;
                        }
                        break;
                    }
                    case 'insertStart': {
                        break;
                    }
                }
                const newRowId = Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["generateRowId"])(tableroViewStore);
                Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["insertNewRow"])(tableroDocumentStore, newRowId, indexToInsertDM);
                for (let col of this.state.columns) {
                    await Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["setCellValue"])(tableroDocumentStore, newRowId, col.id, Object(_dataModel_CellDataUtils__WEBPACK_IMPORTED_MODULE_6__["createDefaultCellData"])(col.type));
                }
                const rowToInsert = {
                    id: newRowId
                };
                Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["setRowViewProperties"])(tableroViewStore, Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_5__["addElementAtIndex"])(orderedRowIds, indexToInsertVM, rowToInsert));
                if (arg.type === 'insertRelative' && arg.newPresenceLocation !== undefined) {
                    const newPresenceColumnId = arg.newPresenceLocation.id;
                    // tslint:disable-next-line: no-floating-promises
                    this.onLocalPresenceValueProposal({
                        type: 'cell',
                        columnId: newPresenceColumnId,
                        rowId: newRowId
                    });
                }
            }, _telemetry__WEBPACK_IMPORTED_MODULE_13__["UserAction"].InsertTableRow, _dataModel_TableroStore__WEBPACK_IMPORTED_MODULE_11__["EditMode"].Edit, undefined /* setNoop */, requestOrigin);
        };
        this.onInsertColumnProposal = (arg, requestOrigin) => {
            this.store.postTransaction(async () => {
                const { tableroDocumentStore, tableroViewStore } = this.store;
                const orderedColumns = Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getColumnViewProperties"])(tableroViewStore);
                let indexToInsert = 0;
                switch (arg.type) {
                    case 'insertRelative': {
                        // Update the ordered columns
                        const indexOfCol = tableroDocumentStore.colIdToIndex[arg.id];
                        if (indexOfCol !== undefined) {
                            indexToInsert = arg.insertPosition === _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__["RelativeInsertPosition"].After ? indexOfCol + 1 : indexOfCol;
                        }
                        break;
                    }
                    case 'insertStart':
                        indexToInsert = 0;
                        break;
                }
                // Currently defining the column data type based on Tablero Configuration Type.
                // TODO: Do we want allow our users to select column type while inserting new column ?
                const configurationType = Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getComponentConfigurationType"])(this.store.tableroViewStore, this.props.model.logger);
                let defaultDataTypeForColumnInsertion;
                if (configurationType === _configuration_Configuration__WEBPACK_IMPORTED_MODULE_23__["Configuration"].richTextTablero) {
                    defaultDataTypeForColumnInsertion = _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__["DataType"].RichText;
                }
                else {
                    defaultDataTypeForColumnInsertion = _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__["DataType"].Text;
                }
                const newColumnId = Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["generateColumnId"])(tableroViewStore);
                Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["insertNewColumn"])(tableroDocumentStore, newColumnId, indexToInsert, defaultDataTypeForColumnInsertion);
                const newDefaultCellData = Object(_dataModel_CellDataUtils__WEBPACK_IMPORTED_MODULE_6__["createDefaultCellData"])(defaultDataTypeForColumnInsertion);
                for (let row of this.state.rows) {
                    await Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["setCellValue"])(tableroDocumentStore, row.id, newColumnId, newDefaultCellData);
                }
                const columnToInsert = {
                    id: newColumnId,
                    width: _StylingConstants__WEBPACK_IMPORTED_MODULE_22__["defaultColumnWidth"]
                };
                Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["setColumnViewProperties"])(tableroViewStore, Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_5__["addElementAtIndex"])(orderedColumns, indexToInsert, columnToInsert));
            }, _telemetry__WEBPACK_IMPORTED_MODULE_13__["UserAction"].InsertTableColumn, _dataModel_TableroStore__WEBPACK_IMPORTED_MODULE_11__["EditMode"].Edit, undefined /* setNoop */, requestOrigin);
        };
        this.onDeleteRowProposal = (rowId, requestOrigin) => {
            this.store.postTransaction(() => {
                const { tableroDocumentStore, tableroViewStore } = this.store;
                const orderedRowIds = Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getRowViewProperties"])(tableroViewStore);
                const indexOfRow = tableroDocumentStore.rowIdToIndex[rowId];
                if (indexOfRow !== undefined) {
                    Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["deleteRow"])(tableroDocumentStore, indexOfRow);
                    Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["setRowViewProperties"])(tableroViewStore, Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_5__["removeIndex"])(orderedRowIds, Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getOrderedRowIndexFromId"])(tableroViewStore, rowId)));
                }
                // Considering successful deletion if there is no error in above process.
                // Update local presence for user who deleted it.
                // TODO: Remove this when the presence is not lost on context menu and available in grabber selection.
                this.updatePresencePostRowDeletion(indexOfRow);
            }, _telemetry__WEBPACK_IMPORTED_MODULE_13__["UserAction"].DeleteTableRow, _dataModel_TableroStore__WEBPACK_IMPORTED_MODULE_11__["EditMode"].Edit, true, requestOrigin);
        };
        this.onDeleteColumnProposal = (columnId, requestOrigin) => {
            this.store.postTransaction(() => {
                const { tableroDocumentStore, tableroViewStore } = this.store;
                const orderedColumns = Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getColumnViewProperties"])(tableroViewStore);
                const indexOfCol = tableroDocumentStore.colIdToIndex[columnId];
                if (indexOfCol !== undefined) {
                    Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["deleteColumn"])(tableroDocumentStore, indexOfCol);
                    Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["setColumnViewProperties"])(tableroViewStore, Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_5__["removeIndex"])(orderedColumns, indexOfCol));
                    // Remove Sort Information from View Store if applied on column being deleted.
                    let sortInformation = Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getSortByColumn"])(tableroViewStore);
                    if (sortInformation && sortInformation.sortByColumnId === columnId) {
                        Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["setSortByColumn"])(tableroViewStore, undefined);
                    }
                    // Considering successful deletion if there is no error in above process.
                    // Update local presence for user who deleted it.
                    // TODO: Remove this when the presence is not lost on context menu and available in grabber selection.
                    this.updatePresencePostColumnDeletion(indexOfCol, this.state.columns);
                }
            }, _telemetry__WEBPACK_IMPORTED_MODULE_13__["UserAction"].DeleteTableColumn, _dataModel_TableroStore__WEBPACK_IMPORTED_MODULE_11__["EditMode"].Edit, undefined /* setNoop */, requestOrigin);
        };
        this.onColumnTypeProposal = (columnId, type) => {
            this.store.postTransaction(async () => {
                const { tableroDocumentStore } = this.store;
                let currentColumnType = Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getColumnType"])(tableroDocumentStore, columnId);
                Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["updateColumnData"])(tableroDocumentStore, columnId, { type });
                for (let row of this.state.rows) {
                    if (currentColumnType !== type) {
                        // tslint:disable-next-line: no-floating-promises
                        await Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["setCellValue"])(tableroDocumentStore, row.id, columnId, Object(_dataModel_CellDataUtils__WEBPACK_IMPORTED_MODULE_6__["createDefaultCellData"])(type));
                    }
                }
            }, _telemetry__WEBPACK_IMPORTED_MODULE_13__["UserAction"].ChangeTableColumnType, _dataModel_TableroStore__WEBPACK_IMPORTED_MODULE_11__["EditMode"].Edit, true /* setNoop */);
        };
        this.onTableTitleProposal = (newTableTitle) => {
            this.store.postTransaction(() => {
                Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["setComponentFriendlyName"])(this.store.tableroViewStore, newTableTitle);
            }, _telemetry__WEBPACK_IMPORTED_MODULE_13__["UserAction"].ChangeTableTitle, _dataModel_TableroStore__WEBPACK_IMPORTED_MODULE_11__["EditMode"].Edit);
        };
        this.onFilterRowsProposal = (columnId, filter) => {
            this.store.postTransaction(() => {
                Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["updateColumnViewProperties"])(this.store.tableroViewStore, columnId, { filter });
            }, _telemetry__WEBPACK_IMPORTED_MODULE_13__["UserAction"].FilterTableRows, _dataModel_TableroStore__WEBPACK_IMPORTED_MODULE_11__["EditMode"].Edit);
        };
        this.onColumnTitleProposal = (columnId, newTitle) => {
            this.store.postTransaction(() => {
                const tableroStore = this.store;
                Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["updateColumnData"])(tableroStore.tableroDocumentStore, columnId, {
                    title: newTitle
                });
            }, _telemetry__WEBPACK_IMPORTED_MODULE_13__["UserAction"].ChangeTableColumnTitle, _dataModel_TableroStore__WEBPACK_IMPORTED_MODULE_11__["EditMode"].Edit, true /* setNoop */);
        };
        this.getPossibleFilterValuesForColumn = (columnId) => {
            return Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getPossibleFilterValuesForColumn"])(this.store, columnId, this.props.model.componentContext);
        };
        this.onReorderColumnProposal = (columnIdToMove, relativeColumnId, insertPosition) => {
            this.store.postTransaction(() => {
                const columnViewData = Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getColumnViewProperties"])(this.store.tableroViewStore);
                const columnToMove = columnViewData.find((c) => c.id === columnIdToMove);
                if (columnToMove !== undefined) {
                    const orderedColumnsWithoutTargetColumn = columnViewData.filter((c) => c.id !== columnIdToMove);
                    const relativeColumnIndex = orderedColumnsWithoutTargetColumn.findIndex((c) => c.id === relativeColumnId);
                    if (relativeColumnIndex >= 0) {
                        const insertionIndex = insertPosition === _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__["RelativeInsertPosition"].After ? relativeColumnIndex + 1 : relativeColumnIndex;
                        const newOrderedColumns = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_5__["addElementAtIndex"])(orderedColumnsWithoutTargetColumn, insertionIndex, columnToMove);
                        Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["setColumnViewProperties"])(this.store.tableroViewStore, newOrderedColumns);
                    }
                }
            }, _telemetry__WEBPACK_IMPORTED_MODULE_13__["UserAction"].ReorderTableColumns, _dataModel_TableroStore__WEBPACK_IMPORTED_MODULE_11__["EditMode"].Edit);
        };
        // Check if the caret is entering into a fluid component cell then call it's setSelection method
        //  for component to manage it's own entry of ip selection.
        this.onMoveSelectionToFluidComponentCell = async (newPresenceLocation, selectionParams) => {
            var _a;
            if (newPresenceLocation.type === 'cell') {
                const component = await Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getCellData"])(this.store.tableroDocumentStore, newPresenceLocation.rowId, newPresenceLocation.columnId, Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getColumnType"])(this.store.tableroDocumentStore, newPresenceLocation.columnId), this.props.model.componentContext, this.props.model.logger);
                (_a = component.ComponentSelection) === null || _a === void 0 ? void 0 : _a.setSelection(selectionParams);
            }
        };
        this.onMoveSelectionToCell = async (newPresenceLocation, selectionParams) => {
            // Set moveSelectionToCellDetails if there is direction else unset it.
            this.moveSelectionToCellDetails = selectionParams;
            if (newPresenceLocation && selectionParams) {
                await this.onMoveSelectionToFluidComponentCell(newPresenceLocation, selectionParams);
            }
            this.setAttributionData(true, newPresenceLocation);
        };
        this.clearSelection = async (currentPresenceLocation, newPresenceLocation) => {
            if (newPresenceLocation && currentPresenceLocation && currentPresenceLocation.type === 'cell') {
                const component = await Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getCellData"])(this.store.tableroDocumentStore, currentPresenceLocation.rowId, currentPresenceLocation.columnId, Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getColumnType"])(this.store.tableroDocumentStore, currentPresenceLocation.columnId), this.props.model.componentContext, this.props.model.logger);
                const componentSelection = component === null || component === void 0 ? void 0 : component.ComponentSelection;
                if (componentSelection === null || componentSelection === void 0 ? void 0 : componentSelection.clearSelection) {
                    componentSelection.clearSelection();
                }
            }
        };
        this.onLocalPresenceChanged = async (currentPresenceLocation, newPresenceLocation, selectionParams) => {
            // tslint:disable-next-line: no-floating-promises
            this.clearSelection(currentPresenceLocation, newPresenceLocation);
            await this.onMoveSelectionToCell(newPresenceLocation, selectionParams);
        };
        this.onLocalPresenceValueProposal = async (newPresenceLocation, selectionParams // Implies we are trying to enter in edit mode in a direction.
        ) => {
            const currentPresenceLocation = this.state.localPresence;
            // Ignore the same proposed presence location.
            if ((currentPresenceLocation === undefined && newPresenceLocation === undefined) ||
                (currentPresenceLocation !== undefined && Object(_presence_Presence__WEBPACK_IMPORTED_MODULE_7__["arePresenceLocationsEqual"])(currentPresenceLocation, newPresenceLocation))) {
                return;
            }
            await this.onLocalPresenceChanged(currentPresenceLocation, newPresenceLocation, selectionParams);
            this.broadcastLocalPresence(newPresenceLocation);
            this.setState({ localPresence: newPresenceLocation });
        };
        this.setCellSelection = (rowId, columnId) => {
            // tslint:disable-next-line: no-floating-promises
            this.onLocalPresenceValueProposal({ type: 'cell', rowId, columnId });
        };
        this.getBasicUserInfo = (clientId) => {
            const { componentContext, presenceManager } = this.props.model;
            if (presenceManager !== undefined) {
                const coauthorInfo = presenceManager.getCoauthorInformationForClient(clientId);
                // Don't show presence when the user is undefined.
                if (!coauthorInfo) {
                    return undefined;
                }
                const userInfo = {
                    userId: coauthorInfo.id,
                    name: coauthorInfo.name,
                    color: coauthorInfo.color
                };
                return userInfo;
            }
            const fluidUser = Object(_dataModel_TableroDocumentUtil__WEBPACK_IMPORTED_MODULE_26__["getFluidUserWithDefaults"])(componentContext, clientId);
            if (fluidUser === undefined) {
                return undefined;
            }
            const userInfo = {
                userId: fluidUser.id,
                name: fluidUser.name
            };
            return userInfo;
        };
        this.registerCommandFactories = () => {
            const insertionCommandGroup = [];
            if (!this.getPreset(_configuration_Presets__WEBPACK_IMPORTED_MODULE_8__["PresetName"].DisableInsertColumns)) {
                insertionCommandGroup.push(this.insertColumnOnLeftCommandFactory);
                insertionCommandGroup.push(this.insertColumnOnRightCommandFactory);
            }
            if (!this.getPreset(_configuration_Presets__WEBPACK_IMPORTED_MODULE_8__["PresetName"].DisableInsertRows)) {
                insertionCommandGroup.push(this.insertRowAboveCommandFactory);
                insertionCommandGroup.push(this.insertRowBelowCommandFactory);
            }
            const deletionCommandGroup = [];
            !this.getPreset(_configuration_Presets__WEBPACK_IMPORTED_MODULE_8__["PresetName"].DisableDeleteRows) && deletionCommandGroup.push(this.deleteRowCommandFactory);
            !this.getPreset(_configuration_Presets__WEBPACK_IMPORTED_MODULE_8__["PresetName"].DisableDeleteColumns) && deletionCommandGroup.push(this.deleteColumnCommandFactory);
            const sortCommandGroup = [];
            !this.getPreset(_configuration_Presets__WEBPACK_IMPORTED_MODULE_8__["PresetName"].DisableSort) && sortCommandGroup.push(this.sortCommandFactory);
            const resizeCommandGroup = [];
            !this.getPreset(_configuration_Presets__WEBPACK_IMPORTED_MODULE_8__["PresetName"].DisableResizeColumnFromCmdMenu) &&
                Object(_TableAppFeatures__WEBPACK_IMPORTED_MODULE_27__["resizeColumnFromCmdMenuFeatureEnabled"])() &&
                resizeCommandGroup.push(this.resizeColumnCommandFactory);
            this.commandFactories.push(sortCommandGroup, insertionCommandGroup, deletionCommandGroup, resizeCommandGroup);
        };
        this.sortCommandFactory = (commandContext) => {
            const columnId = commandContext.columnId;
            let sortDirection;
            if (columnId === undefined || !(commandContext.rowId === undefined)) {
                return undefined;
            }
            this.state.columns.map((column) => {
                if (column.id === commandContext.columnId) {
                    sortDirection = column.sortDirection;
                }
            });
            let sortIcon;
            let sortDisplayName;
            let sortId;
            let newSortDirection;
            if (sortDirection === _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__["SortDirection"].Ascending) {
                sortIcon = 'FFXCSortUp';
                sortDisplayName = _StringTable__WEBPACK_IMPORTED_MODULE_24__["columnMenuSortLabel"] + ' ' + _StringTable__WEBPACK_IMPORTED_MODULE_24__["sortDescending"];
                sortId = _StringTable__WEBPACK_IMPORTED_MODULE_24__["sortDescending"];
                newSortDirection = _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__["SortDirection"].Descending;
            }
            else if (sortDirection === _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__["SortDirection"].Descending) {
                sortIcon = 'FFXCSortDown';
                sortDisplayName = _StringTable__WEBPACK_IMPORTED_MODULE_24__["columnMenuSortLabel"] + ' ' + _StringTable__WEBPACK_IMPORTED_MODULE_24__["sortAscending"];
                sortId = _StringTable__WEBPACK_IMPORTED_MODULE_24__["sortAscending"];
                newSortDirection = _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__["SortDirection"].Ascending;
            }
            else {
                sortIcon = 'FFXCSort';
                sortDisplayName = _StringTable__WEBPACK_IMPORTED_MODULE_24__["columnMenuSortLabel"];
                sortId = _StringTable__WEBPACK_IMPORTED_MODULE_24__["columnMenuSortLabel"];
                newSortDirection = _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__["SortDirection"].Ascending;
            }
            const sortCommand = {
                icon: sortIcon,
                displayName: sortDisplayName,
                id: sortId,
                disabled: this.state.isReadOnlyMode,
                execute: () => {
                    this.onSortByColumnProposal(columnId, newSortDirection);
                }
            };
            return sortCommand;
        };
        this.insertRowAboveCommandFactory = (commandContext) => {
            let insertArgument;
            if (commandContext.columnId === undefined || commandContext.rowId === undefined) {
                // Check for commandContext.rowId is to hide the option for 'inserting a row above' in heading row.
                return undefined;
            }
            if (commandContext.rowId !== undefined) {
                insertArgument = {
                    type: 'insertRelative',
                    id: commandContext.rowId,
                    insertPosition: _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__["RelativeInsertPosition"].Before
                };
            }
            const insertRowAboveCommand = {
                icon: Object(_fluidx_icons__WEBPACK_IMPORTED_MODULE_28__["componentsIconName"])('FFXCInsertRowsAbove'),
                displayName: this.getAboveRowLabel(),
                id: 'insertRowAbove',
                disabled: this.state.isReadOnlyMode,
                execute: () => {
                    this.onInsertRowProposal(insertArgument);
                }
            };
            return insertRowAboveCommand;
        };
        this.insertRowBelowCommandFactory = (commandContext) => {
            let insertArgument;
            if (commandContext.columnId === undefined) {
                return undefined;
            }
            if (commandContext.rowId === undefined) {
                if (commandContext.isHeaderCell !== undefined && commandContext.isHeaderCell) {
                    // we have to special-case the header row to make sure that we can still insert rows
                    // when the header row is right-clicked. Since the header row isn't a true row in the
                    // sparseMatrix (and therefore doesn't have an id), we check to see if we've been invoked
                    // in a header cell, and then insert a row at the start.
                    insertArgument = { type: 'insertStart' };
                }
                else {
                    return undefined;
                }
            }
            else {
                insertArgument = {
                    type: 'insertRelative',
                    id: commandContext.rowId,
                    insertPosition: _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__["RelativeInsertPosition"].After
                };
            }
            const insertRowCommand = {
                icon: Object(_fluidx_icons__WEBPACK_IMPORTED_MODULE_28__["componentsIconName"])('FFXCInsertRowsBelow'),
                displayName: this.getBelowRowLabel(),
                id: 'insertRowBelow',
                disabled: this.state.isReadOnlyMode,
                execute: () => {
                    this.onInsertRowProposal(insertArgument);
                }
            };
            return insertRowCommand;
        };
        this.insertColumnOnLeftCommandFactory = (commandContext) => {
            if (commandContext.columnId === undefined) {
                return undefined;
            }
            const insertArgument = {
                type: 'insertRelative',
                id: commandContext.columnId,
                insertPosition: _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__["RelativeInsertPosition"].Before
            };
            const insertColumnCommand = {
                icon: Object(_fluidx_icons__WEBPACK_IMPORTED_MODULE_28__["componentsIconName"])('FFXCInsertColumnsLeft'),
                displayName: 'Insert column on left',
                id: 'insertColumnLeft',
                disabled: this.state.isReadOnlyMode,
                execute: () => {
                    this.onInsertColumnProposal(insertArgument);
                }
            };
            return insertColumnCommand;
        };
        this.insertColumnOnRightCommandFactory = (commandContext) => {
            if (commandContext.columnId === undefined) {
                return undefined;
            }
            const insertArgument = {
                type: 'insertRelative',
                id: commandContext.columnId,
                insertPosition: _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__["RelativeInsertPosition"].After
            };
            const insertColumnCommand = {
                icon: Object(_fluidx_icons__WEBPACK_IMPORTED_MODULE_28__["componentsIconName"])('FFXCInsertColumnsRight'),
                displayName: 'Insert column on right',
                id: 'insertColumnRight',
                disabled: this.state.isReadOnlyMode,
                execute: () => {
                    this.onInsertColumnProposal(insertArgument);
                }
            };
            return insertColumnCommand;
        };
        this.deleteRowCommandFactory = (commandContext) => {
            const rowId = commandContext.rowId;
            if (rowId === undefined || commandContext.columnId === undefined) {
                return undefined;
            }
            const deleteRowCommand = {
                icon: Object(_fluidx_icons__WEBPACK_IMPORTED_MODULE_28__["componentsIconName"])('FFXCDeleteRows'),
                displayName: this.getDeleteRowLabel(),
                id: 'deleteRow',
                disabled: this.state.isReadOnlyMode,
                execute: () => {
                    this.onDeleteRowProposal(rowId);
                }
            };
            return deleteRowCommand;
        };
        this.deleteColumnCommandFactory = (commandContext) => {
            const columnId = commandContext.columnId;
            if (columnId === undefined) {
                return undefined;
            }
            const deleteColumnCommand = {
                icon: Object(_fluidx_icons__WEBPACK_IMPORTED_MODULE_28__["componentsIconName"])('FFXCDeleteColumns'),
                displayName: 'Delete column',
                id: 'deleteColumn',
                disabled: this.state.isReadOnlyMode,
                execute: () => {
                    this.onDeleteColumnProposal(columnId);
                }
            };
            return deleteColumnCommand;
        };
        this.resizeColumnCommandFactory = (commandContext) => {
            const rowId = commandContext.rowId;
            const columnId = commandContext.columnId;
            if (columnId === undefined || (rowId === undefined && !commandContext.isHeaderCell)) {
                return undefined;
            }
            const resizeColumnCommand = {
                icon: undefined,
                displayName: _StringTable__WEBPACK_IMPORTED_MODULE_24__["resizeColumnLabel"],
                id: 'resizeColumn',
                disabled: this.state.isReadOnlyMode,
                execute: () => {
                    this.onResizeColumnCmd(rowId, columnId);
                }
            };
            return resizeColumnCommand;
        };
        this.getOwnerDocument = () => {
            if (!this.wrapperRef.current || !this.wrapperRef.current.ownerDocument) {
                const { logger } = this.props.model;
                logger &&
                    Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_15__["sendErrorEvent"])(logger, {
                        eventName: _telemetry__WEBPACK_IMPORTED_MODULE_12__["ErrorEvent"].CommandingSurfaceError,
                        errorCode: _telemetry__WEBPACK_IMPORTED_MODULE_12__["ErrorEvent"].CommandingSurfaceError
                    });
                return undefined;
            }
            return this.wrapperRef.current.ownerDocument;
        };
        this.onInvokeContextMenu = (commandContext) => {
            const commandGroups = Object(_commanding_getCommands__WEBPACK_IMPORTED_MODULE_10__["getCommandingSurfaceCommands"])(commandContext, this.commandFactories);
            if (commandGroups.length > 0) {
                // Only show the commanding surface if we have commands to show.
                const ownerDocument = this.getOwnerDocument();
                if (ownerDocument) {
                    const commandingSurfaceData = {
                        clientX: commandContext.event.clientX,
                        clientY: commandContext.event.clientY,
                        ownerDocument,
                        specialCommands: {},
                        commandGroups,
                        triggerByKeyboardEvent: commandContext.event.button === 0,
                        onDismiss: commandContext.onDismissCallback
                    };
                    this.props.model.sendCommandingSurfaceData(commandingSurfaceData);
                }
            }
        };
        this.getPreset = (presetName) => {
            const configurationType = Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getComponentConfigurationType"])(this.store.tableroViewStore, this.props.model.logger);
            return Object(_configuration_ConfigurationUtils__WEBPACK_IMPORTED_MODULE_9__["getConfigurationPreset"])(configurationType, presetName);
        };
        this.onUndoProposal = () => {
            const store = this.store;
            store.postTransaction(() => {
                var _a;
                (_a = store.undoHandler) === null || _a === void 0 ? void 0 : _a.undo();
            }, undefined, _dataModel_TableroStore__WEBPACK_IMPORTED_MODULE_11__["EditMode"].Undo);
        };
        this.onRedoProposal = () => {
            const store = this.store;
            store.postTransaction(() => {
                var _a;
                (_a = store.undoHandler) === null || _a === void 0 ? void 0 : _a.redo();
            }, undefined, _dataModel_TableroStore__WEBPACK_IMPORTED_MODULE_11__["EditMode"].Redo);
        };
        this.onPresenceSignalReceived = (message) => {
            if (this.props.model.presenceManager !== undefined) {
                this.onPresenceSignalReceivedNew(message);
                return;
            }
            const presenceLocation = message.content;
            const localClientId = this.props.model.componentContext.clientId;
            // If presence signal is for the local client, ignore it
            if (message.clientId === localClientId) {
                return;
            }
            // tslint:disable-next-line: no-floating-promises
            this.updateCoauthorsMap(message.clientId, presenceLocation);
        };
        this.onPresenceSignalReceivedNew = (message) => {
            const { presenceManager } = this.props.model;
            // If the signal was sent from a client using container presence manager, and current client is using container presence manager too,
            // ignore this signal because our container presence manager will handle that client's presence.
            if (message.content !== undefined && message.content.forLegacyClients) {
                return;
            }
            // Otherwise, the signal was sent from an old client which doesn't have container presence manager.
            // Let container presence manager handle it.
            if (presenceManager) {
                const presenceLocation = message.content;
                presenceManager.processLegacyClientPresence(message.clientId, this.props.model.componentId, presenceLocation, undefined);
            }
        };
        this.onRequestPresenceSignalReceived = (message) => {
            if (this.props.model.presenceManager !== undefined) {
                this.onRequestPresenceSignalReceivedNew(message);
                return;
            }
            this.broadcastLocalPresence(this.state.localPresence);
        };
        this.onRequestPresenceSignalReceivedNew = (message) => {
            // If the signal was sent from a client using container presence manager, and current client is using container presence manager too,
            // ignore this signal because our container presence manager will handle that client's presence.
            if (message.content !== undefined && message.content.forLegacyClients) {
                return;
            }
            // Otherwise, an old client is asking for my presence.
            this.broadcastLocalPresenceNew(this.state.localPresence);
        };
        this.onPresenceChanged = (clientId, componentId, selection, _additionalProps) => {
            // Ignore presence in other components.
            if (componentId !== this.props.model.componentId) {
                return;
            }
            const presenceLocation = selection;
            // tslint:disable-next-line: no-floating-promises
            this.updateCoauthorsMap(clientId, presenceLocation);
        };
        this.onPresenceActiveStatusChanged = (clientInfo, isActive) => {
            const { componentId } = this.props.model;
            if (isActive) {
                // Switching from inactive to active: show presence. Presence data is obtained from containerPresenceManager.
                const presenceLocation = clientInfo && clientInfo.presence && clientInfo.presence.selectionsMap[componentId];
                if (presenceLocation) {
                    // tslint:disable-next-line: no-floating-promises
                    this.updateCoauthorsMap(clientInfo.id, presenceLocation);
                }
            }
            else {
                // Switching from active to inactive: remove presence.
                // tslint:disable-next-line: no-floating-promises
                this.updateCoauthorsMap(clientInfo.id, undefined);
            }
        };
        this.onMoveSelectionToHost = (selectionParams) => {
            const { onMoveSelectionToHost } = this.props.model;
            this.setAttributionData(false);
            return onMoveSelectionToHost(selectionParams);
        };
        this.onSetIsFocusInTable = (focusState) => {
            const { setIsFocusInTable } = this.props.model;
            this.setAttributionData(focusState);
            return setIsFocusInTable(focusState);
        };
        this.createAttributionUXData = (presenceLocation) => {
            const { rows, columns } = this.state;
            const { componentContext } = this.props.model;
            const tableroStore = this.store;
            const fluidUser = Object(_dataModel_TableroDocumentUtil__WEBPACK_IMPORTED_MODULE_26__["getFluidUserWithDefaults"])(componentContext, componentContext.clientId);
            if (fluidUser === undefined) {
                return undefined;
            }
            const attributionDataWithAllyText = Object(_utilities_attributionUtils__WEBPACK_IMPORTED_MODULE_25__["getAttributionDataWithAllyText"])(tableroStore, rows, columns, fluidUser, presenceLocation);
            const renderUserHoverHighlight = (selectedUserData) => {
                const selectedUserRange = selectedUserData.range;
                if (selectedUserRange && selectedUserRange.type === 'cells') {
                    this.setState({ hoverHighlightAttributionMap: selectedUserRange.cells });
                }
            };
            const renderUserClickHighlight = (selectedUserData) => {
                const selectedUserRange = selectedUserData.range;
                if (selectedUserRange && selectedUserRange.type === 'cells') {
                    this.setState({ clickHighlightAttributionMap: selectedUserRange.cells });
                }
            };
            const onCalloutDismiss = (clearHoverHighlight, clearClickHighlight) => {
                if (clearHoverHighlight) {
                    this.setState({ hoverHighlightAttributionMap: undefined });
                }
                if (clearClickHighlight) {
                    this.setState({ clickHighlightAttributionMap: undefined });
                }
                if (clearClickHighlight && clearHoverHighlight) {
                    this.setAttributionData(false);
                    if (presenceLocation && presenceLocation.type === 'cell') {
                        this.setCellSelection(presenceLocation.rowId, presenceLocation.columnId);
                    }
                    else {
                        this.setCellSelection(rows[1].id, columns[0].id);
                    }
                }
            };
            return {
                attributionData: attributionDataWithAllyText.attributionData,
                onItemClick: renderUserClickHighlight,
                onItemHover: renderUserHoverHighlight,
                onCalloutDismiss,
                getAccessibilityDivText: () => {
                    return attributionDataWithAllyText.accessibilityText;
                }
            };
        };
        this.onNestedComponentDataChangeCallback = async () => {
            // This will trigger the notifyRootMapChanged task listener which will push the event to task
            // TODO: Remove this noop and fix it
            const { componentContext } = this.props.model;
            const tableroStore = this.store;
            if (this.tableroDocument === undefined) {
                this.tableroDocument = await this.getTableroDocument(tableroStore, componentContext);
            }
            this.tableroDocument.notifyTableUpdate();
        };
        this.store = this.props.model.tableroStore;
        const tableroHeaderData = Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getHeaderDataForRender"])(this.store);
        this.state = Object.assign(Object.assign({}, tableroHeaderData), { rows: [], coauthorsMap: new Map(), localPresence: undefined, clickHighlightAttributionMap: undefined, hoverHighlightAttributionMap: undefined, isReadOnlyMode: this.props.model.componentRuntime.deltaManager.readonly || false, resizeColumnProperties: { rowId: undefined, columnId: undefined } });
        this.wrapperRef = this.props.model.wrapperRef;
        this.stylingService = this.props.model.stylingService;
        this.moveSelectionToCellDetails = undefined;
        const { presenceManager } = this.props.model;
        if (presenceManager !== undefined) {
            presenceManager.on(_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["PresenceManagerEvent"].clientSelectionChanged, this.onPresenceChanged);
            presenceManager.on(_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_1__["PresenceManagerEvent"].clientActiveStatusChanged, this.onPresenceActiveStatusChanged);
        }
    }
    componentDidMount() {
        this.registerEventListeners();
        this.registerCommandFactories();
        if (this.wrapperRef) {
            Object(_utilities_stylingUtils__WEBPACK_IMPORTED_MODULE_20__["listenOnThemeChange"])(this.stylingService, this.themeChangeListener);
            // Update theme variables by default for first time when component is rendered
            Object(_utilities_stylingUtils__WEBPACK_IMPORTED_MODULE_20__["clearAndUpdateThemeVariables"])(undefined, this.wrapperRef.current, this.stylingService);
        }
        const { componentContext, logger } = this.props.model;
        const tableroStore = this.store;
        if (tableroStore.tableroViewStore.rootMap.get('componentConfigurationType') === _configuration_Configuration__WEBPACK_IMPORTED_MODULE_23__["Configuration"].tasks) {
            // tslint:disable-next-line: no-floating-promises
            this.notifyDeleteTableFromView(tableroStore, componentContext, false);
        }
        const tracker = _fluidx_utilities__WEBPACK_IMPORTED_MODULE_18__["ActivityTracker"].start(_telemetry__WEBPACK_IMPORTED_MODULE_14__["Activity"].GetTableBodyData, logger);
        Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getBodyDataForRender"])(tableroStore, componentContext, logger)
            .then((tableroBodyData) => {
            this.setState(tableroBodyData);
            tracker.setResult(true);
        })
            .catch((error) => {
            tracker.setResult(false, {}, error);
        });
    }
    async notifyDeleteTableFromView(tableroStore, componentContext, isDeleted) {
        if (this.tableroDocument === undefined) {
            this.tableroDocument = await this.getTableroDocument(tableroStore, componentContext);
        }
        this.tableroDocument.notifyDeleteTableFromView(isDeleted);
    }
    componentWillUnmount() {
        const { componentContext, presenceManager, componentRuntime } = this.props.model;
        const tableroStore = this.store;
        const { tableroViewStore, tableroDocumentStore } = tableroStore;
        Object(_utilities_stylingUtils__WEBPACK_IMPORTED_MODULE_20__["removeThemeEventListener"])(this.stylingService, this.themeChangeListener);
        componentRuntime.deltaManager.off(readonly, this.onReadOnlyEvent);
        tableroViewStore.rootMap.off(valueChanged, this.viewStoreRootMapValueChangedListener);
        tableroDocumentStore.rootMap.off(valueChanged, this.documentStoreRootMapValueChangedListener);
        tableroDocumentStore.table.removeListener(sequenceDelta, this.onReceivedSequenceDeltaOp);
        if (tableroDocumentStore.table.sequenceDeltaListened !== undefined) {
            tableroDocumentStore.table.sequenceDeltaListened -= 1;
        }
        if (tableroStore.tableroViewStore.rootMap.get('componentConfigurationType') === _configuration_Configuration__WEBPACK_IMPORTED_MODULE_23__["Configuration"].tasks) {
            // tslint:disable-next-line: no-floating-promises
            this.notifyDeleteTableFromView(tableroStore, componentContext, true);
        }
        delete tableroDocumentStore.table.sequenceDeltaListened;
        if (presenceManager === undefined) {
            // Remove the local user from the presence map.
            // TODO: TABLERO: This code doesn't always run on scenarios like closing a browser tab. We should find out the right way to do this in Fluid.
            this.broadcastLocalPresence(undefined);
            componentContext.getAudience().removeListener(removeMember, this.removeMemberForPresence);
        }
        window.removeEventListener(beforeunload, this.onWindowBeforeUnload);
    }
    async getTableroDocument(tableroStore, componentContext) {
        const tableroDocumentId = tableroStore.tableroViewStore.rootMap.get('tableroDocumentId');
        let tableroDocument = undefined;
        let response = await componentContext.containerRuntime.request({ url: tableroDocumentId });
        if (response.status === 200 && response.mimeType === 'fluid/component') {
            tableroDocument = response.value;
        }
        return tableroDocument;
    }
    async onDocumentStoreRootMapValueChanged() {
        const { componentContext, logger } = this.props.model;
        const { rows, columns, localPresence } = this.state;
        const tableroStore = this.store;
        if (tableroStore.tableroDocumentStore.tableTitle &&
            tableroStore.tableroViewStore.rootMap.get(componentFriendlyName) !== tableroStore.tableroDocumentStore.tableTitle) {
            tableroStore.tableroViewStore.rootMap.set(componentFriendlyName, tableroStore.tableroDocumentStore.tableTitle);
        }
        // Do not re-render during undo. Instead we should only re-render AFTER undo.
        // The problem is when we edit the table, there are a few operations have rootMap change in the middle of the edits,
        // not after all the edits (not noop change, but some other useful changes). Since root map change triggers re-rendering,
        // so when we undo all the ops in the reverse order, it's not guaranteed it'll be in a consistent state when we undo the
        // root map changes.
        // The current solution is let's remember the edit state and do not re-render if we are still undoing things but only re-render
        // when all the undo is finished which consistency is guaranteed.
        // The whole reason we need this check is related how Tablero triggers re-render, and I think the logic of re-rendering should be
        // revisit so that we can get rid of these kind of tricks.
        if (!tableroStore.revertInProgress) {
            const tableroHeaderData = Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getHeaderDataForRender"])(tableroStore);
            let tableroBodyData;
            const tracker = _fluidx_utilities__WEBPACK_IMPORTED_MODULE_18__["ActivityTracker"].start(_telemetry__WEBPACK_IMPORTED_MODULE_14__["Activity"].GetTableBodyData, logger);
            if (tableroStore.tableroViewStore.rootMap.get('componentConfigurationType') === _configuration_Configuration__WEBPACK_IMPORTED_MODULE_23__["Configuration"].tasks) {
                this.updateViewStore(tableroStore);
            }
            try {
                tableroBodyData = await Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getBodyDataForRender"])(tableroStore, componentContext, logger);
                tracker.setResult(true);
            }
            catch (error) {
                tracker.setResult(false, {}, error);
                /**
                 * In case we fail to get updated tablero body data, we should keep both header and body data in sync.
                 * We return here to avoid updating header data state alone
                 *
                 * TODO: Task 3732997: Review tablero async data fetch and boot logic
                 */
                return;
            }
            this.setState(Object.assign(Object.assign({}, tableroHeaderData), tableroBodyData));
            this.updatePresenceOnRowColDeletion(tableroHeaderData, tableroBodyData, rows, columns, localPresence);
        }
    }
    updateViewStore(tableroStore) {
        // Function will be removed after TableroDocument APIs will be available
        let rowId;
        rowId = [];
        let rowMap = new Map();
        const orderedRowIds = Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getRowViewProperties"])(tableroStore.tableroViewStore);
        for (let i = 0; i < orderedRowIds.length; i += 1) {
            rowMap.set(orderedRowIds[i].id, i);
        }
        for (let rowIndex in tableroStore.tableroDocumentStore.rowIdToIndex) {
            if (rowIndex !== undefined)
                rowId.push(rowIndex);
        }
        // New Row inserted in tableroDocument
        if (orderedRowIds.length < rowId.length) {
            for (let i = 0; i < rowId.length; i += 1) {
                if (rowMap.get(rowId[i]) === undefined) {
                    const rowToInsert = {
                        id: rowId[i]
                    };
                    Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["setRowViewProperties"])(tableroStore.tableroViewStore, Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_5__["addElementAtIndex"])(orderedRowIds, tableroStore.tableroDocumentStore.rowIdToIndex[rowId[i]], rowToInsert));
                    const nextRowId = Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["generateTableroDocumentNextRowId"])(tableroStore.tableroDocumentStore);
                    tableroStore.tableroViewStore.rootMap.set('nextRowId', nextRowId);
                }
            }
        }
        // Row deleted in tableroDocument
        if (orderedRowIds.length > rowId.length) {
            for (let i = 0; i < orderedRowIds.length; i += 1) {
                if (tableroStore.tableroDocumentStore.rowIdToIndex[orderedRowIds[i].id] === undefined) {
                    Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["setRowViewProperties"])(tableroStore.tableroViewStore, Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_5__["removeIndex"])(orderedRowIds, i));
                }
            }
        }
    }
    updatePresenceOnRowColDeletion(tableroHeaderData, tableroBodyData, rows, columns, localPresence) {
        /**
         * Check if the user has local presence and recent update caused a row to be deleted.
         * TODO: We can't handle the case of user being offline and when he is back online, more than 1 rows are deleted.
         * Above case will require runtime to provide presence or id of last row deleted.
         *
         * If the row is deleted, we check if the current user was in same row and update + broadcast his/her presence.
         */
        if (localPresence && rows.length - 1 === tableroBodyData.rows.length) {
            const indexOfDeletedRow = this.firstMissingElem(rows, tableroBodyData.rows);
            const deletedRow = rows[indexOfDeletedRow];
            // TODO: deletedRow is coming as undefined.
            // Need to understand exact scenario when indexOfDeletedRow will be -1.
            // Possibly both the users deleting same row at same time or is on a slow n/w.
            if (deletedRow) {
                if (localPresence.rowId === deletedRow.id) {
                    this.updatePresencePostRowDeletion(indexOfDeletedRow);
                }
            }
            else {
                const logger = this.props.model.logger;
                logger &&
                    Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_15__["sendErrorEvent"])(logger, {
                        eventName: _telemetry__WEBPACK_IMPORTED_MODULE_12__["ErrorEvent"].DeletedRowMissingError
                    });
            }
        }
        /**
         * Check if the user has local presence and recent update caused a column to be deleted.
         * TODO: We can't handle the case of user being offline and when he is back online, more than 1 columns are deleted.
         * Above case will require runtime to provide presence or id of last column deleted.
         *
         * If the column is deleted, we check if the current user was in same column and update + broadcast his/her presence.
         */
        if (localPresence && columns.length - 1 === tableroHeaderData.columns.length) {
            const indexOfDeletedColumn = this.firstMissingElem(columns, tableroHeaderData.columns);
            const deletedColumn = columns[indexOfDeletedColumn];
            // TODO: deletedColumn is coming as undefined.
            // Need to understand exact scenario when indexOfDeletedColumn will be -1.
            // Possibly both the users deleting same column at same time or is on a slow n/w.
            if (deletedColumn) {
                if (localPresence.columnsId === deletedColumn.id) {
                    this.updatePresencePostColumnDeletion(indexOfDeletedColumn, columns);
                }
            }
            else {
                const logger = this.props.model.logger;
                logger &&
                    Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_15__["sendErrorEvent"])(logger, {
                        eventName: _telemetry__WEBPACK_IMPORTED_MODULE_12__["ErrorEvent"].DeletedColumnMissingError
                    });
            }
        }
    }
    getVisibleColumn(columnId, direction = _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__["SearchDirection"].Right) {
        const getNextSearchIndex = (index) => (direction === _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__["SearchDirection"].Left ? index - 1 : index + 1);
        const { tableroViewStore } = this.store;
        const columns = Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getColumnViewProperties"])(tableroViewStore);
        const hiddenColumnsIds = Object(_dataModel_TableroAdapter__WEBPACK_IMPORTED_MODULE_4__["getHiddenColumnIds"])(tableroViewStore);
        const columnIndex = columns.map((col) => col.id).indexOf(columnId);
        let nextColumnIndex = getNextSearchIndex(columnIndex);
        while (nextColumnIndex >= 0 && nextColumnIndex < columns.length) {
            if (hiddenColumnsIds.indexOf(columns[nextColumnIndex].id) === -1) {
                return columns[nextColumnIndex];
            }
            nextColumnIndex = getNextSearchIndex(nextColumnIndex);
        }
        return undefined;
    }
    getNextFocusableColumn(columnId) {
        // Scan right for any visible column and then left if not found else return undefined
        return (this.getVisibleColumn(columnId, _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__["SearchDirection"].Right) || this.getVisibleColumn(columnId, _view_table_Table_props__WEBPACK_IMPORTED_MODULE_2__["SearchDirection"].Left));
    }
    setFocusOnNextFocussableColumn(columnId) {
        const nextFocussableColumn = this.getNextFocusableColumn(columnId);
        nextFocussableColumn &&
            this.onLocalPresenceValueProposal({
                type: 'column',
                columnId: nextFocussableColumn.id
            });
    }
    broadcastLocalPresence(presenceLocation) {
        const { logger, presenceManager, componentRuntime } = this.props.model;
        if (presenceManager !== undefined) {
            this.broadcastLocalPresenceNew(presenceLocation);
            return;
        }
        try {
            if (componentRuntime.connected) {
                componentRuntime.submitSignal(presenceSignal, presenceLocation);
            }
        }
        catch (error) {
            logger &&
                Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_15__["sendErrorEvent"])(logger, {
                    eventName: _telemetry__WEBPACK_IMPORTED_MODULE_12__["ErrorEvent"].PresenceSignalSubmissionError,
                    errorCode: _telemetry__WEBPACK_IMPORTED_MODULE_12__["ErrorEvent"].PresenceSignalSubmissionError
                }, error);
        }
    }
    broadcastLocalPresenceNew(presenceLocation) {
        const { logger, presenceManager, componentRuntime } = this.props.model;
        if (presenceManager !== undefined) {
            presenceManager.setComponentSelection(this.props.model.componentId, presenceLocation);
        }
        // For compatibility, we also need to submit a signal from Tablero, so other clients that don't use containerPresenceManager know where we are.
        // TODO: Remove following code when max unsupported version have been bumped to a version that contains containerPresenceManager implementation.
        try {
            if (componentRuntime.connected) {
                let content = presenceLocation !== undefined ? presenceLocation : {};
                content.forLegacyClients = true; // set a flag to indicate this signal is sent from a new client just for compatibility.
                componentRuntime.submitSignal(presenceSignal, content);
            }
        }
        catch (error) {
            logger &&
                Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_15__["sendErrorEvent"])(logger, {
                    eventName: _telemetry__WEBPACK_IMPORTED_MODULE_12__["ErrorEvent"].PresenceSignalSubmissionError,
                    errorCode: _telemetry__WEBPACK_IMPORTED_MODULE_12__["ErrorEvent"].PresenceSignalSubmissionError
                }, error);
        }
    }
    async updateCoauthorsMap(clientId, presenceLocation) {
        var _a;
        const basicUserInfo = this.getBasicUserInfo(clientId);
        // Delete the existing entry for this user
        let updatedCoauthorsMap = Object(_presence_Presence__WEBPACK_IMPORTED_MODULE_7__["deleteUserFromMap"])(this.state.coauthorsMap, basicUserInfo);
        if (!presenceLocation || !basicUserInfo) {
            this.setState({ coauthorsMap: updatedCoauthorsMap });
            return;
        }
        const finishUpdatingCoauthorsMap = (userInfo) => {
            const rowId = presenceLocation.type === 'cell' ? presenceLocation.rowId : _Constants__WEBPACK_IMPORTED_MODULE_21__["headerRowIndex"];
            // Add the user to the map at the new location.
            // Since this is a pure component, we should be treating the coauthorMap as immutable.
            // It is safe to add a user in-place to the map since deleteUserFromMap creates a new coauthorMap above.
            Object(_presence_Presence__WEBPACK_IMPORTED_MODULE_7__["addUserToMapInPlace"])(updatedCoauthorsMap, rowId, presenceLocation.columnId, userInfo);
            this.setState({ coauthorsMap: updatedCoauthorsMap });
        };
        if (basicUserInfo.color) {
            finishUpdatingCoauthorsMap(basicUserInfo);
            return;
        }
        // Otherwise, if have a user but don't yet have its color info yet, try to fetch it:
        const color = await Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_19__["resolvePresenceColor"])(this.props.model.presenceColorProvider, (_a = basicUserInfo.userId) !== null && _a !== void 0 ? _a : clientId /* use clientId in case userId is undefined or empty */, this.props.model.logger);
        finishUpdatingCoauthorsMap(Object.assign(Object.assign({}, basicUserInfo), { color }));
    }
    setAttributionData(focusState, presenceLocation) {
        const { setAttributionUXDataCallback } = this.props.model;
        if (setAttributionUXDataCallback) {
            if (focusState) {
                setAttributionUXDataCallback(this.createAttributionUXData(presenceLocation));
            }
            else {
                // Clear previously applied highlight
                setAttributionUXDataCallback(undefined);
                this.setState({ clickHighlightAttributionMap: undefined, hoverHighlightAttributionMap: undefined });
            }
        }
    }
    getBelowRowLabel() {
        switch (this.props.model.tableroconfig) {
            case _configuration_Configuration__WEBPACK_IMPORTED_MODULE_23__["Configuration"].tasks:
                return _StringTable__WEBPACK_IMPORTED_MODULE_24__["insertTaskBelowLabel"];
            default:
                return _StringTable__WEBPACK_IMPORTED_MODULE_24__["insertRowBelowLabel"];
        }
    }
    getAboveRowLabel() {
        switch (this.props.model.tableroconfig) {
            case _configuration_Configuration__WEBPACK_IMPORTED_MODULE_23__["Configuration"].tasks:
                return _StringTable__WEBPACK_IMPORTED_MODULE_24__["insertTaskAboveLabel"];
            default:
                return _StringTable__WEBPACK_IMPORTED_MODULE_24__["insertRowAboveLabel"];
        }
    }
    getDeleteRowLabel() {
        switch (this.props.model.tableroconfig) {
            case _configuration_Configuration__WEBPACK_IMPORTED_MODULE_23__["Configuration"].tasks:
                return _StringTable__WEBPACK_IMPORTED_MODULE_24__["deleteTaskLabel"];
            default:
                return _StringTable__WEBPACK_IMPORTED_MODULE_24__["deleteRowLabel"];
        }
    }
    render() {
        const { componentContext, logger, sendInsertComponentRequestData, getSelectionEnterDetails, tableroconfig } = this.props.model;
        const { rows, columns, title, coauthorsMap, localPresence, hasHiddenColumns, clickHighlightAttributionMap, hoverHighlightAttributionMap, isReadOnlyMode, resizeColumnProperties } = this.state;
        return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_view_table_Table__WEBPACK_IMPORTED_MODULE_3__["Table"], { componentContext: componentContext, enableUndoRedo: this.props.model.enableUndoRedo, logger: logger, columns: columns, rows: rows, title: title, onUndoProposal: this.onUndoProposal, onRedoProposal: this.onRedoProposal, onCellValueProposal: this.onCellValueProposal, onColumnResizeProposal: this.onColumnResizeProposal, onColumnTypeProposal: this.onColumnTypeProposal, onSortByColumnProposal: this.onSortByColumnProposal, onInsertRowProposal: this.onInsertRowProposal, onInsertColumnProposal: this.onInsertColumnProposal, onHideColumnProposal: this.onFilterColumnProposal, onShowAllColumnsProposal: this.onShowAllColumnsProposal, hasHiddenColumns: hasHiddenColumns, onTableTitleProposal: this.onTableTitleProposal, onFilterRowsProposal: this.onFilterRowsProposal, getPossibleFilterValuesForColumn: this.getPossibleFilterValuesForColumn, onReorderColumnProposal: this.onReorderColumnProposal, onColumnTitleProposal: this.onColumnTitleProposal, onLocalPresenceValueProposal: this.onLocalPresenceValueProposal, localPresenceLocation: localPresence, coauthorsMap: coauthorsMap, wrapperRef: this.wrapperRef, onInvokeContextMenu: this.onInvokeContextMenu, onMoveSelectionToHost: this.onMoveSelectionToHost, getSelectionEnterDetails: getSelectionEnterDetails, setIsFocusInTable: this.onSetIsFocusInTable, getPreset: this.getPreset, sendInsertComponentRequestData: sendInsertComponentRequestData, stylingService: this.stylingService, moveSelectionToCellDetails: this.moveSelectionToCellDetails, tableroconfig: tableroconfig, clickHighlightAttributionMap: clickHighlightAttributionMap, hoverHighlightAttributionMap: hoverHighlightAttributionMap, onDeleteRowProposal: this.onDeleteRowProposal, onDeleteColumnProposal: this.onDeleteColumnProposal, onNestedComponentDataChangeCallback: this.onNestedComponentDataChangeCallback, isReadOnlyMode: isReadOnlyMode, onResizeColumnLeaveCmd: this.onResizeColumnCmd, resizeColumnProperties: resizeColumnProperties, onCellSelectionProposal: this.setCellSelection }));
    }
};
// window.FluidTable = FluidTable;
// cspell:ignore Callout


/***/ }),

/***/ "../tablero/lib/view/table/GrabberTableStyles.js":
/*!*******************************************************!*\
  !*** ../tablero/lib/view/table/GrabberTableStyles.js ***!
  \*******************************************************/
/*! exports provided: columnGrabberTableClass, rowGrabberWrapperClass, rowGrabberTableClass, rowGrabberHeaderClass */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "columnGrabberTableClass", function() { return columnGrabberTableClass; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "rowGrabberWrapperClass", function() { return rowGrabberWrapperClass; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "rowGrabberTableClass", function() { return rowGrabberTableClass; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "rowGrabberHeaderClass", function() { return rowGrabberHeaderClass; });
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/memoize.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/styling/7.13.5_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/styling/lib/index.js");
/* harmony import */ var _StylingConstants__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../StylingConstants */ "../tablero/lib/view/StylingConstants.js");


const columnGrabberTableClass = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["memoizeFunction"])(() => Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["mergeStyles"])({
    tableLayout: 'fixed',
    width: '0px',
    height: '0',
    borderCollapse: 'separate !important',
    selectors: {
        [office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]]: {
            MsHighContrastAdjust: 'none'
        }
    }
}));
const rowGrabberWrapperClass = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["memoizeFunction"])(() => Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["mergeStyles"])({
    position: 'absolute',
    left: `-${_StylingConstants__WEBPACK_IMPORTED_MODULE_2__["rowGrabberWidth"] + _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["cellWrapperBorderWidth"]}px`
}));
const rowGrabberTableClass = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["memoizeFunction"])(() => Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["mergeStyles"])({
    tableLayout: 'fixed',
    width: '0px',
    height: '0',
    borderCollapse: 'separate',
    selectors: {
        [office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]]: {
            MsHighContrastAdjust: 'none'
        }
    }
}));
const rowGrabberHeaderClass = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["memoizeFunction"])((headerHeight) => Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["mergeStyles"])({
    height: headerHeight,
    width: _StylingConstants__WEBPACK_IMPORTED_MODULE_2__["rowGrabberWidth"],
    padding: 'unset'
}));


/***/ }),

/***/ "../tablero/lib/view/table/RowGrabbersTable.js":
/*!*****************************************************!*\
  !*** ../tablero/lib/view/table/RowGrabbersTable.js ***!
  \*****************************************************/
/*! exports provided: RowGrabbersTable */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "RowGrabbersTable", function() { return RowGrabbersTable; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _GrabberTableStyles__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./GrabberTableStyles */ "../tablero/lib/view/table/GrabberTableStyles.js");
/* harmony import */ var _Constants__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../Constants */ "../tablero/lib/view/Constants.js");
/* harmony import */ var _cells_GrabberCell__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../cells/GrabberCell */ "../tablero/lib/view/cells/GrabberCell.js");
/* harmony import */ var _TableDefinitions__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./TableDefinitions */ "../tablero/lib/view/table/TableDefinitions.js");
/* harmony import */ var _TableStyle__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ./TableStyle */ "../tablero/lib/view/table/TableStyle.js");






/**
 * RowGrabbersTable is a table on the left of data table, which has a row grabber for each data table row
 */
const RowGrabbersTable = react__WEBPACK_IMPORTED_MODULE_0__["memo"](function RowGrabbersTable(props) {
    const { rows, rowHeights, onInsertRowProposal, onDeleteRowProposal, onCellStyleOverride, hoveredTableCellRowId, getLastKnownLocalPresence, onLocalPresenceValueProposal, activateDragPreview, selectedRowId, onRowColumnSelect, isReadOnlyMode } = props;
    const onGrabberMouseDown = (id) => {
        // User may drag grabber, which should result in
        // row re-order preview.
        activateDragPreview(_TableDefinitions__WEBPACK_IMPORTED_MODULE_4__["TableElementType"].Row, id);
    };
    /**
     * Invoked when user clicks on grabber to select it. Clears grabber selection if id is undefined.
     */
    const onSelectRow = (id) => {
        onRowColumnSelect(id, _TableDefinitions__WEBPACK_IMPORTED_MODULE_4__["TableElementType"].Row);
    };
    return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { className: Object(_GrabberTableStyles__WEBPACK_IMPORTED_MODULE_1__["rowGrabberWrapperClass"])() },
        react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("table", { className: `${Object(_GrabberTableStyles__WEBPACK_IMPORTED_MODULE_1__["rowGrabberTableClass"])()} ${_TableStyle__WEBPACK_IMPORTED_MODULE_5__["tableBorderSpacingClassName"]}` },
            react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("thead", null,
                react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("tr", null,
                    react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("th", { className: Object(_GrabberTableStyles__WEBPACK_IMPORTED_MODULE_1__["rowGrabberHeaderClass"])(rowHeights.get(_Constants__WEBPACK_IMPORTED_MODULE_2__["headerRowIndex"])) }))),
            react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("tbody", null, rows.map((row) => (react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("tr", { key: row.id },
                react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_cells_GrabberCell__WEBPACK_IMPORTED_MODULE_3__["GrabberCell"], { type: 0 /* Row */, rowHeight: rowHeights.get(row.id), id: row.id, onInsertProposal: onInsertRowProposal, onDeleteProposal: onDeleteRowProposal, onCellStyleOverride: onCellStyleOverride, hoveredTableCellId: hoveredTableCellRowId, getLastKnownLocalPresence: getLastKnownLocalPresence, onLocalPresenceValueProposal: onLocalPresenceValueProposal, isSelected: row.id === selectedRowId, onMouseDown: onGrabberMouseDown, onSelect: onSelectRow, isReadOnlyMode: isReadOnlyMode }))))))));
});


/***/ }),

/***/ "../tablero/lib/view/table/Table.js":
/*!******************************************!*\
  !*** ../tablero/lib/view/table/Table.js ***!
  \******************************************/
/*! exports provided: Table */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Table", function() { return Table; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var uuid_v4__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! uuid/v4 */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/uuid/3.4.0/node_modules/uuid/v4.js");
/* harmony import */ var uuid_v4__WEBPACK_IMPORTED_MODULE_1___default = /*#__PURE__*/__webpack_require__.n(uuid_v4__WEBPACK_IMPORTED_MODULE_1__);
/* harmony import */ var _Table_props__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./Table.props */ "../tablero/lib/view/table/Table.props.js");
/* harmony import */ var _columns_ColumnHeader__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../columns/ColumnHeader */ "../tablero/lib/view/columns/ColumnHeader.js");
/* harmony import */ var _rows_TableRow__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../rows/TableRow */ "../tablero/lib/view/rows/TableRow.js");
/* harmony import */ var _header_TableHeader__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../header/TableHeader */ "../tablero/lib/view/header/TableHeader.js");
/* harmony import */ var _TableKeyboarding__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ./TableKeyboarding */ "../tablero/lib/view/table/TableKeyboarding.js");
/* harmony import */ var _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! @fluidx/office-fluid-types */ "../office-fluid-types/lib/ComponentContracts/ComponentSelection.js");
/* harmony import */ var _StylingConstants__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../StylingConstants */ "../tablero/lib/view/StylingConstants.js");
/* harmony import */ var _ms_office_bohemia_ux__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! @ms/office-bohemia-ux */ "../office-bohemia-ux/lib/utilities/nodeContainsTarget.js");
/* harmony import */ var _presence_Presence__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ../presence/Presence */ "../tablero/lib/view/presence/Presence.js");
/* harmony import */ var _Constants__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(/*! ../Constants */ "../tablero/lib/view/Constants.js");
/* harmony import */ var _utils_Helper__WEBPACK_IMPORTED_MODULE_12__ = __webpack_require__(/*! ../utils/Helper */ "../tablero/lib/view/utils/Helper.js");
/* harmony import */ var _rows_AddRow__WEBPACK_IMPORTED_MODULE_13__ = __webpack_require__(/*! ../rows/AddRow */ "../tablero/lib/view/rows/AddRow.js");
/* harmony import */ var _StringTable__WEBPACK_IMPORTED_MODULE_14__ = __webpack_require__(/*! ../StringTable */ "../tablero/lib/view/StringTable.js");
/* harmony import */ var _TableGradient__WEBPACK_IMPORTED_MODULE_15__ = __webpack_require__(/*! ./TableGradient */ "../tablero/lib/view/table/TableGradient.js");
/* harmony import */ var _utils_ResizeObserverWrapper__WEBPACK_IMPORTED_MODULE_16__ = __webpack_require__(/*! ./../utils/ResizeObserverWrapper */ "../tablero/lib/view/utils/ResizeObserverWrapper.js");
/* harmony import */ var _cells_CellDefinitions__WEBPACK_IMPORTED_MODULE_17__ = __webpack_require__(/*! ../cells/CellDefinitions */ "../tablero/lib/view/cells/CellDefinitions.js");
/* harmony import */ var _cells_CellHelpers__WEBPACK_IMPORTED_MODULE_18__ = __webpack_require__(/*! ../cells/CellHelpers */ "../tablero/lib/view/cells/CellHelpers.js");
/* harmony import */ var _TableDefinitions__WEBPACK_IMPORTED_MODULE_19__ = __webpack_require__(/*! ./TableDefinitions */ "../tablero/lib/view/table/TableDefinitions.js");
/* harmony import */ var _ColumnGrabbersTable__WEBPACK_IMPORTED_MODULE_20__ = __webpack_require__(/*! ./ColumnGrabbersTable */ "../tablero/lib/view/table/ColumnGrabbersTable.js");
/* harmony import */ var _RowGrabbersTable__WEBPACK_IMPORTED_MODULE_21__ = __webpack_require__(/*! ./RowGrabbersTable */ "../tablero/lib/view/table/RowGrabbersTable.js");
/* harmony import */ var _TableAppFeatures__WEBPACK_IMPORTED_MODULE_22__ = __webpack_require__(/*! ./TableAppFeatures */ "../tablero/lib/view/table/TableAppFeatures.js");
/* harmony import */ var _TableStyle__WEBPACK_IMPORTED_MODULE_23__ = __webpack_require__(/*! ./TableStyle */ "../tablero/lib/view/table/TableStyle.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_24__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/KeyCodes.js");
/* harmony import */ var _configuration_Presets__WEBPACK_IMPORTED_MODULE_25__ = __webpack_require__(/*! ../../configuration/Presets */ "../tablero/lib/configuration/Presets.js");


























// The DragManager is lazily loaded to improve performance.
const DragManager = react__WEBPACK_IMPORTED_MODULE_0__["lazy"](() => __webpack_require__.e(/*! import() | DragManager */ "DragManager").then(__webpack_require__.bind(null, /*! ../dragAndDrop/DragManager */ "../tablero/lib/view/dragAndDrop/DragManager.js")).then((module) => ({
    default: module.DragManager
})));
const tabIndex = -1;
const taskNameColumnIndex = 1;
/**
 * Handler for mouse down event on tableRef.
 * @param ev MouseEvent
 */
const onTableMouseDown = (ev) => {
    // Mouse handler of tableWrapperRef calls ev.preventDefault() on mouse down. That causes
    // issues with mouse interaction on text cell. So, all tableRef mouse down events are
    // stopped here.
    ev.stopPropagation();
};
/**
 * Handler for mouse down event on tableWrapperRef.
 */
const onTableElementWrapperMouseDown = (ev) => {
    // Bug 4188683: Can't scroll horizontally when Table height exceeds available viewport.
    // This prevents wrapperRef from taking focus when mouse down event is fired on tableWrapperRef.
    // This handler is called when scrollbar is mouse clicked and dragged. We don't want wrapperRef
    // to take focus in such case.
    ev.preventDefault();
};
const ariaTableRole = 'table';
const ariaRowRole = 'row';
const Table = react__WEBPACK_IMPORTED_MODULE_0__["memo"](function Table(props) {
    const { componentContext, logger, wrapperRef, columns, rows, title, localPresenceLocation, coauthorsMap, onSortByColumnProposal, onCellValueProposal, onColumnResizeProposal, onTableTitleProposal, onInsertRowProposal, onDeleteRowProposal, onInsertColumnProposal, onDeleteColumnProposal, onColumnTitleProposal, onLocalPresenceValueProposal, onInvokeContextMenu, onMoveSelectionToHost, getSelectionEnterDetails, setIsFocusInTable, getPreset, enableUndoRedo, onUndoProposal, onRedoProposal, sendInsertComponentRequestData, moveSelectionToCellDetails, onNestedComponentDataChangeCallback, isReadOnlyMode, onResizeColumnLeaveCmd, resizeColumnProperties, onCellSelectionProposal, clickHighlightAttributionMap, hoverHighlightAttributionMap } = props;
    const ariaMessageDivRef = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](null);
    // A unique ID for every view instance.
    // useRef is used to prevent re-calculation of uuid with every edit on Table title, click on cell, etc
    const tableroUuid = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](uuid_v4__WEBPACK_IMPORTED_MODULE_1__());
    const tableTitleId = 'tableTitle-' + tableroUuid.current;
    // Build Demo Hack: We are hiding the last column spacer for the Build Demo, this should actually be the 2 * columns.length + 1
    const rowSpacerColSpan = 2 * columns.length; // There is a spacer for every column, plus one before the first column.
    const setColumnSelection = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((columnId) => {
        onLocalPresenceValueProposal({ type: 'column', columnId });
    }, []);
    const tableRef = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](null);
    const tableWrapperRef = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](null);
    const tableTitleRef = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](null);
    const addRowBtnRef = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](null);
    const isAddRowBtnEnabled = getPreset(_configuration_Presets__WEBPACK_IMPORTED_MODULE_25__["PresetName"].EnableAddRowButton);
    const [hoveredTableCellIds, setHoveredTableCellIds] = react__WEBPACK_IMPORTED_MODULE_0__["useState"]({
        rowId: undefined,
        columnId: undefined
    });
    /**
     * Callback function that is invoked when there is mouse enter or mouse leave event from a cell
     */
    const onTableCellHover = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((columnId, rowId) => {
        setHoveredTableCellIds({ columnId, rowId });
    }, []);
    const [cellStyleOverride, setCellStyleOverride] = react__WEBPACK_IMPORTED_MODULE_0__["useState"]();
    /**
     * Callback function that is invoked when cell style override is provided
     */
    const onCellStyleOverride = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((args) => {
        let overrideAllEdges = false;
        let newAllEdgesValue;
        if (args.allEdgesOverride !== undefined &&
            args.allEdgesOverride !== _cells_CellDefinitions__WEBPACK_IMPORTED_MODULE_17__["CellStyleOverrideOption"].MaintainOldOverride) {
            overrideAllEdges = true;
            if (args.allEdgesOverride === _cells_CellDefinitions__WEBPACK_IMPORTED_MODULE_17__["CellStyleOverrideOption"].Default) {
                newAllEdgesValue = undefined;
            }
            else {
                newAllEdgesValue = args.allEdgesOverride;
            }
        }
        let overrideSingleEdge = false;
        let newSingleEdgeValue;
        if (args.singleEdgeOverride !== undefined &&
            args.singleEdgeOverride !== _cells_CellDefinitions__WEBPACK_IMPORTED_MODULE_17__["CellStyleOverrideOption"].MaintainOldOverride) {
            overrideSingleEdge = true;
            if (args.singleEdgeOverride === _cells_CellDefinitions__WEBPACK_IMPORTED_MODULE_17__["CellStyleOverrideOption"].Default) {
                newSingleEdgeValue = undefined;
            }
            else {
                newSingleEdgeValue = args.singleEdgeOverride;
            }
        }
        let overrideOpacity = false;
        let newOpacityValue;
        if (args.opacityOverride !== undefined && args.opacityOverride !== _cells_CellDefinitions__WEBPACK_IMPORTED_MODULE_17__["CellStyleOverrideOption"].MaintainOldOverride) {
            overrideOpacity = true;
            if (args.opacityOverride === _cells_CellDefinitions__WEBPACK_IMPORTED_MODULE_17__["CellStyleOverrideOption"].Default) {
                newOpacityValue = undefined;
            }
            else {
                newOpacityValue = args.opacityOverride;
            }
        }
        setCellStyleOverride((prevState) => ({
            allEdges: overrideAllEdges ? newAllEdgesValue : prevState === null || prevState === void 0 ? void 0 : prevState.allEdges,
            singleEdge: overrideSingleEdge ? newSingleEdgeValue : prevState === null || prevState === void 0 ? void 0 : prevState.singleEdge,
            opacity: overrideOpacity ? newOpacityValue : prevState === null || prevState === void 0 ? void 0 : prevState.opacity
        }));
    }, []);
    // Moves the scroll bar if needed when we arrow key through table cells
    const updateTableScrollOnTableKeyboarding = (target, newColIndex, tableScrollDirection) => {
        if (!tableWrapperRef.current) {
            return;
        }
        if (newColIndex === 0) {
            // If the new cell we navigated to is in the first column move the scroll bar to the beginning
            tableWrapperRef.current.scrollLeft = 0;
            return;
        }
        // Column width of the column that has the cell we navigated to
        const newColWidth = columns[newColIndex].width;
        // The table cell we arrow key on to move to the next/previous cell
        const targetElement = target;
        const tableWrapperBoundingRect = tableWrapperRef.current.getBoundingClientRect();
        const eventTargetBoundingRect = targetElement.getBoundingClientRect();
        switch (tableScrollDirection) {
            // Determine if the scroll bar needs to be moved then move by the same amount as the column width we are navigating to in order to bring the column into view
            case _Table_props__WEBPACK_IMPORTED_MODULE_2__["TableScrollDirection"].Left: {
                // If we are navigating through table cells to the left and if the previous column cell beginning is before the table wrapper beginning then the scroll should move
                if (eventTargetBoundingRect.left - newColWidth < tableWrapperBoundingRect.left) {
                    tableWrapperRef.current.scrollLeft -= newColWidth;
                }
                break;
            }
            case _Table_props__WEBPACK_IMPORTED_MODULE_2__["TableScrollDirection"].Right: {
                // If we are navigating through table cells to the right and if the next column cell ending is after the table wrapper ending then the scroll should move
                if (eventTargetBoundingRect.left + eventTargetBoundingRect.width + newColWidth >
                    tableWrapperBoundingRect.left + tableWrapperBoundingRect.width) {
                    tableWrapperRef.current.scrollLeft += newColWidth;
                }
                break;
            }
        }
    };
    const getTableBounds = () => {
        if (tableRef.current) {
            return tableRef.current.getBoundingClientRect();
        }
        return undefined;
    };
    const onMoveToTableTitle = () => {
        if (tableTitleRef.current) {
            tableTitleRef.current.focus();
        }
    };
    const onMoveToTableHeader = () => {
        if (columns.length > 0) {
            setColumnSelection(columns[0].id);
        }
    };
    const getCaretHorizontalPosition = () => {
        const tableBounds = getTableBounds();
        const tableLeftOffset = tableBounds ? tableBounds.left : 0;
        let horizontalOffset = tableLeftOffset + _StylingConstants__WEBPACK_IMPORTED_MODULE_8__["cellWrapperPaddingLeftAndRight"] + _StylingConstants__WEBPACK_IMPORTED_MODULE_8__["cellWrapperBorderWidth"];
        if (localPresenceLocation !== undefined) {
            const columnId = localPresenceLocation.columnId;
            const currentColumnIndex = columns.findIndex((col) => col.id === columnId);
            for (let index = 0; index < currentColumnIndex; index += 1) {
                horizontalOffset += columns[index].width + _StylingConstants__WEBPACK_IMPORTED_MODULE_8__["cellWrapperBorderWidth"];
            }
        }
        return horizontalOffset;
    };
    /**
     * This function will be invoked from Table Keyboarding to pass focus on 'Add Row' button
     * As we leave table, we are updating Local Presence Value here.
     */
    const focusOnAddRowButton = () => {
        var _a;
        onLocalPresenceValueProposal(undefined);
        if (isAddRowBtnEnabled) {
            (_a = addRowBtnRef === null || addRowBtnRef === void 0 ? void 0 : addRowBtnRef.current) === null || _a === void 0 ? void 0 : _a.focus();
        }
    };
    // This callback function is used by other components to give away control
    // to tablero when they can't handle the keyboarding on it's own.
    const onMoveSelectionFromCellCallback = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((localPresenceLocation) => {
        return Object(_TableKeyboarding__WEBPACK_IMPORTED_MODULE_6__["getTableKeyboardHandler"])(localPresenceLocation, onLocalPresenceValueProposal, rows, columns, onMoveSelectionToHost, updateTableScrollOnTableKeyboarding, !isReadOnlyMode ? onInsertRowProposal : undefined, getCaretHorizontalPosition, onMoveToTableTitle, enableUndoRedo ? onUndoProposal : undefined, enableUndoRedo ? onRedoProposal : undefined, isAddRowBtnEnabled ? focusOnAddRowButton : undefined);
    }, [rows, columns, updateTableScrollOnTableKeyboarding, getCaretHorizontalPosition]);
    // Create a reference to the callback function.
    const onMoveSelectionFromCellCallbackRef = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](onMoveSelectionFromCellCallback);
    // Store the refreshed callback function.
    onMoveSelectionFromCellCallbackRef.current = onMoveSelectionFromCellCallback;
    // Set the onCellCaretLeave to point to updated callback reference.
    const onMoveSelectionFromCell = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((localPresenceLocation) => {
        return onMoveSelectionFromCellCallbackRef.current(localPresenceLocation);
    }, []);
    const onTableKeyDown = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"](Object(_TableKeyboarding__WEBPACK_IMPORTED_MODULE_6__["getTableKeyboardHandler"])(localPresenceLocation, onLocalPresenceValueProposal, rows, columns, onMoveSelectionToHost, updateTableScrollOnTableKeyboarding, !isReadOnlyMode ? onInsertRowProposal : undefined, getCaretHorizontalPosition, onMoveToTableTitle, enableUndoRedo ? onUndoProposal : undefined, enableUndoRedo ? onRedoProposal : undefined, isAddRowBtnEnabled ? focusOnAddRowButton : undefined), [localPresenceLocation, rows, columns, onLocalPresenceValueProposal]);
    const onTableBlur = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((ev) => {
        if (tableWrapperRef.current !== null) {
            // If the table element loses focus, we should clear the local presence.
            // We should make sure we don't clear presence when the element receiving focus (ev.relatedTarget) is also within the table.
            // We are also considering elements from floating UI originated from hosted components to be part of the table.
            // Using tableWrapperRef to include the Grabbers in nodeContainsTarget method.
            // This allows us to keep the selection in that cell.
            const isSelectionInTableWrapperBody = ev.relatedTarget === null ||
                (ev.relatedTarget instanceof Node && !Object(_ms_office_bohemia_ux__WEBPACK_IMPORTED_MODULE_9__["nodeContainsTarget"])(tableWrapperRef.current, ev.relatedTarget));
            if (isSelectionInTableWrapperBody) {
                onLocalPresenceValueProposal(undefined);
                if (ev.relatedTarget !== tableTitleRef.current) {
                    // If focus is not in title or table body
                    setIsFocusInTable(false);
                }
            }
        }
    }, [tableWrapperRef]);
    const selectLastCell = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"](() => {
        if (rows.length > 0 && columns.length > 0) {
            const rowId = rows[rows.length - 1].id;
            const colId = columns[columns.length - 1].id;
            onCellSelectionProposal(rowId, colId);
        }
    }, [onCellSelectionProposal, rows, columns]);
    const selectCellFromHorizontalCaretPosition = (caretX, rowIndex) => {
        var _a;
        const tableRect = (_a = tableRef.current) === null || _a === void 0 ? void 0 : _a.getBoundingClientRect();
        if (tableRect && rows.length > 0 && columns.length > 0) {
            if (caretX > tableRect.right) {
                selectLastCell();
            }
            else if (tableRect.left < caretX) {
                // loop to find the closest cell to the given caret position
                let currentCellStart = tableRect.left;
                for (let index = 0; index < columns.length; index += 1) {
                    if (currentCellStart + columns[index].width > caretX) {
                        onCellSelectionProposal(rows[rowIndex].id, columns[index].id);
                        break;
                    }
                    currentCellStart = currentCellStart + columns[index].width;
                }
            }
            else {
                onCellSelectionProposal(rows[rowIndex].id, columns[0].id);
            }
        }
    };
    const onTableFocus = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((ev) => {
        var _a, _b;
        // If the focus event is coming from the table wrapper it means focus is coming from outside the table
        // check the caret direction to put focus in the corresponding place
        if (ev.target === wrapperRef.current) {
            const selectionEnterDetails = getSelectionEnterDetails();
            if (selectionEnterDetails !== undefined) {
                switch (selectionEnterDetails.direction) {
                    case _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_7__["SelectionDirection"].left: {
                        if (isAddRowBtnEnabled) {
                            focusOnAddRowButton();
                        }
                        else {
                            selectLastCell();
                        }
                        break;
                    }
                    case _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_7__["SelectionDirection"].up: {
                        if (isAddRowBtnEnabled) {
                            focusOnAddRowButton();
                        }
                        else {
                            if ((_a = selectionEnterDetails.currentCoordinates) === null || _a === void 0 ? void 0 : _a.x) {
                                selectCellFromHorizontalCaretPosition((_b = selectionEnterDetails.currentCoordinates) === null || _b === void 0 ? void 0 : _b.x, rows.length - 1);
                            }
                            else {
                                selectLastCell();
                            }
                        }
                        break;
                    }
                    case _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_7__["SelectionDirection"].right:
                    case _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_7__["SelectionDirection"].down: {
                        onMoveToTableHeader();
                        break;
                    }
                }
            }
            else {
                // attempt to focus the first focusable element
                onMoveToTableTitle();
            }
        }
        setIsFocusInTable(true);
    }, [onMoveToTableHeader, selectLastCell, wrapperRef]);
    const onTableContextMenu = (event) => {
        // Don't show the browser's context menu in any location in Tablero.
        event.preventDefault();
        // TODO: What commands to show when we invoke the context menu from here?
    };
    const [gradientState, setGradientState] = react__WEBPACK_IMPORTED_MODULE_0__["useState"]({
        leftWidth: 0,
        rightWidth: 0,
        height: 0,
        offsetTop: 0
    });
    // Used to disable border gradient effect on browsers that don't support ResizeObserver.
    const enableHorizontalScrollBarGradient = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](true);
    const updateHorizontalScrollBarGradient = (element) => {
        if (tableRef.current && enableHorizontalScrollBarGradient.current) {
            // Width of left and right segment that are invisible due to horizontal scroll bar.
            const scrollLeft = element.scrollLeft;
            const scrollRight = element.scrollWidth - element.scrollLeft - element.clientWidth;
            // horizontalScrollBarGradientWidth is default width of gradient div. To create a smooth
            // animation during appearance / disappearance of gradient div, the width of gradient div
            // is set to scrollable length when scrollable length is less than above mentioned default width.
            let newRightGradientWidth = scrollRight > _StylingConstants__WEBPACK_IMPORTED_MODULE_8__["horizontalScrollBarGradientWidth"] ? _StylingConstants__WEBPACK_IMPORTED_MODULE_8__["horizontalScrollBarGradientWidth"] : scrollRight;
            let newLeftGradientWidth = scrollLeft > _StylingConstants__WEBPACK_IMPORTED_MODULE_8__["horizontalScrollBarGradientWidth"] ? _StylingConstants__WEBPACK_IMPORTED_MODULE_8__["horizontalScrollBarGradientWidth"] : scrollLeft;
            // Total width is not wide enough to accommodate gradient divs. The effect is disabled in such cases.
            if (newLeftGradientWidth + newRightGradientWidth > element.clientWidth) {
                newLeftGradientWidth = 0;
                newRightGradientWidth = 0;
            }
            let gradientHeight = tableRef.current.getBoundingClientRect().height;
            let gradientOffsetTop = tableRef.current.offsetTop;
            setGradientState({
                leftWidth: newLeftGradientWidth,
                rightWidth: newRightGradientWidth,
                height: gradientHeight,
                offsetTop: gradientOffsetTop
            });
        }
    };
    const onTableWrapperScroll = (e) => {
        updateHorizontalScrollBarGradient(e.currentTarget);
    };
    const [rowHeights, setRowHeights] = react__WEBPACK_IMPORTED_MODULE_0__["useState"](new Map());
    const headerRowRef = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](null);
    const updateRowHeights = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((rowId, rowHeight) => {
        setRowHeights((prevRowHeights) => {
            let updatedRowHeights = new Map(prevRowHeights);
            updatedRowHeights.set(rowId, rowHeight);
            return updatedRowHeights;
        });
    }, []);
    /**
     * Attaching ResizeObserver for syncing the height of data table header row with RowGrabbersTable header row
     */
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        // tslint:disable-next-line: no-floating-promises
        Object(_utils_ResizeObserverWrapper__WEBPACK_IMPORTED_MODULE_16__["ensureResizeObserver"])().then(() => {
            if (!headerRowRef.current) {
                return;
            }
            const resizeObserver = new _utils_ResizeObserverWrapper__WEBPACK_IMPORTED_MODULE_16__["default"](() => {
                if (headerRowRef.current) {
                    updateRowHeights(_Constants__WEBPACK_IMPORTED_MODULE_11__["headerRowIndex"], headerRowRef.current.getBoundingClientRect().height);
                }
            });
            resizeObserver.observe(headerRowRef.current);
            return () => {
                // Remove all Elements from this observer
                resizeObserver.disconnect();
            };
        });
    }, [headerRowRef]);
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        // tslint:disable-next-line: no-floating-promises
        Object(_utils_ResizeObserverWrapper__WEBPACK_IMPORTED_MODULE_16__["ensureResizeObserver"])()
            .then(() => {
            const resizeObserver = new _utils_ResizeObserverWrapper__WEBPACK_IMPORTED_MODULE_16__["default"]((entries) => {
                entries.forEach((entry) => {
                    const element = entry.target;
                    if (element.className === Object(_TableStyle__WEBPACK_IMPORTED_MODULE_23__["tableElementWrapperClassName"])()) {
                        updateHorizontalScrollBarGradient(element);
                    }
                    else if (element.className === _TableStyle__WEBPACK_IMPORTED_MODULE_23__["tableClassName"] && element.parentElement) {
                        updateHorizontalScrollBarGradient(element.parentElement);
                    }
                });
            });
            // Changes is width of table-wrapper <div> and <table> are observed for calculation of height and width of TableGradient.
            // table-wrapper <div> width can change when user resizes browser window, affecting horizontal scroll state of table.
            if (tableWrapperRef.current)
                resizeObserver.observe(tableWrapperRef.current);
            // <table> width will change when user changes width of a column, affecting horizontal scroll state of table.
            if (tableRef.current)
                resizeObserver.observe(tableRef.current);
            return () => {
                if (tableWrapperRef.current)
                    resizeObserver.unobserve(tableWrapperRef.current);
                if (tableRef.current)
                    resizeObserver.unobserve(tableRef.current);
            };
        })
            .catch(() => {
            enableHorizontalScrollBarGradient.current = false;
        });
    }, []);
    const ariaMessage = (msg) => {
        if (ariaMessageDivRef.current) {
            ariaMessageDivRef.current.innerText = msg;
        }
    };
    const getColumnsForAttribution = (rowId, attributionMap) => {
        if (attributionMap) {
            return attributionMap[rowId];
        }
        return undefined;
    };
    /**
     * Returns cell presence location details
     * @param columnId
     * @param rowId
     */
    const getLocalPresenceLocation = (columnId, rowId) => {
        // Only the row or column that has local presence needs to know about it, others can receive undefined.
        if (columnId) {
            return Object(_presence_Presence__WEBPACK_IMPORTED_MODULE_10__["isPresenceInColumn"])(localPresenceLocation, columnId) ? localPresenceLocation : undefined;
        }
        if (rowId) {
            return Object(_presence_Presence__WEBPACK_IMPORTED_MODULE_10__["isPresenceInRow"])(localPresenceLocation, rowId) ? localPresenceLocation : undefined;
        }
        return undefined;
    };
    // Indicates if a row / column is selected. If so, then type (row / column) and its id. 'undefined' id
    // means no selection.
    const [rowColumnSelectionState, setRowColumnSelectionState] = react__WEBPACK_IMPORTED_MODULE_0__["useState"]({
        id: undefined,
        type: _TableDefinitions__WEBPACK_IMPORTED_MODULE_19__["TableElementType"].Column
    });
    /**
     * Changes RowColumnSelectionState to reflect new row/column selection. Will clear selection if id is undefined.
     */
    const onRowColumnSelect = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((id, type) => {
        setRowColumnSelectionState({ id, type });
    }, []);
    // Save the local user's last known presence in table cells
    const lastKnownLocalPresence = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](undefined);
    react__WEBPACK_IMPORTED_MODULE_0__["useEffect"](() => {
        if (localPresenceLocation) {
            lastKnownLocalPresence.current = localPresenceLocation;
        }
    }, [localPresenceLocation]);
    // Callback function that returns the local user's last known presence in table cells
    const getLastKnownLocalPresence = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"](() => {
        return lastKnownLocalPresence.current;
    }, [lastKnownLocalPresence]);
    /**
     * dragAction is passed to <DragPreview> as prop where it is implemented.
     * It is used to forward grabber mouse down event to <DragManager> synchronously.
     * <DragPreview> must receive this event synchronously because <DragPreview> needs
     * to attach mouse up event to document. Missing mouse up event may result in
     * unexpected errors.
     */
    const dragAction = react__WEBPACK_IMPORTED_MODULE_0__["useRef"](null);
    /**
     * This method forwards all mouse down event to TableDragPreview component.
     *
     * @param tableElementType Type of dragPreview that will start.
     * @param id Id of row / column that was clicked.
     */
    const activateDragPreview = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((tableElementType, id) => {
        if (dragAction.current) {
            dragAction.current.activateDragPreview(tableElementType, id);
        }
    }, [dragAction.current]);
    // API to access <TableCellWrapper>.
    const tableCellApi = react__WEBPACK_IMPORTED_MODULE_0__["useRef"]({});
    // Gets a cloned preview of <TableCellWrapper>.
    const getTableCellPreview = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((columnId, rowId) => {
        return tableCellApi.current[Object(_cells_CellHelpers__WEBPACK_IMPORTED_MODULE_18__["getCellId"])(columnId, rowId)].getPreview();
    }, [tableCellApi.current]);
    /**
     * Function to get left and right coordinates for all columns.
     */
    const getColumnBounds = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"](() => {
        const columnBounds = {};
        if (tableRef.current && tableRef.current.tHead && tableRef.current.tHead.rows.length !== 0) {
            const columnHeaders = tableRef.current.tHead.rows[0].children;
            // Index starts from 1 as first cell (<td>) in each <tr> is used by row grabber.
            for (let i = 0; i < columnHeaders.length; i += 1) {
                const clientRect = columnHeaders[i].getBoundingClientRect();
                // (i-1) because row grabber cell is irrelevant to columnBounds.
                columnBounds[columns[i].id] = { start: clientRect.left, end: clientRect.right };
            }
        }
        return columnBounds;
    }, [columns]);
    /**
     * Function to get top and bottom coordinates for all rows.
     */
    const getRowBounds = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"](() => {
        const rowBounds = {};
        if (tableRef.current) {
            const htmlRows = tableRef.current.rows;
            // Index starts from 1 as first <tr> (table row) is used by column headers.
            for (let i = 1; i < htmlRows.length; i += 1) {
                const clientRect = htmlRows[i].getBoundingClientRect();
                // (i-1) because rowBounds don't need bounds for column header.
                rowBounds[rows[i - 1].id] = { start: clientRect.top, end: clientRect.bottom };
            }
        }
        return rowBounds;
    }, [rows]);
    /**
     * Function to get top and bottom coordinates for all rows when tableElementType is row.
     * Function to get left and right coordinates for all columns when tableElementType is column.
     */
    const getRowColumnBounds = react__WEBPACK_IMPORTED_MODULE_0__["useCallback"]((tableElementType) => {
        return tableElementType === _TableDefinitions__WEBPACK_IMPORTED_MODULE_19__["TableElementType"].Row ? getRowBounds() : getColumnBounds();
    }, [getColumnBounds, getRowBounds]);
    /**
     * Function to handle focus when arrow keys are pressed on focused 'AddRow' button
     */
    const setFocusAfterAddRowBtn = (keycode) => {
        if (keycode === office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_24__["KeyCodes"].up) {
            const getLastKnownLocation = getLastKnownLocalPresence();
            if (getLastKnownLocation) {
                onLocalPresenceValueProposal(getLastKnownLocation);
            }
            else {
                // To handle a refreshed page, when we don't know the last cell location
                if (rows.length === 0) {
                    setColumnSelection(columns[1].id);
                }
                else {
                    onCellSelectionProposal(rows[rows.length - 1].id, columns[0].id);
                }
            }
        }
        else if (keycode === office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_24__["KeyCodes"].left) {
            if (rows.length > 0) {
                selectLastCell();
            }
            else {
                setColumnSelection(columns[columns.length - 1].id);
            }
        }
        else if (keycode === office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_24__["KeyCodes"].right) {
            // Move focus to the right in hosting canvas
            onMoveSelectionToHost({ mode: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_7__["SelectionMode"].ip, direction: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_7__["SelectionDirection"].right });
        }
        else if (keycode === office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_24__["KeyCodes"].down) {
            // As arrow down is pressed focus should go to the hosting canvas
            onMoveSelectionToHost({ mode: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_7__["SelectionMode"].ip, direction: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_7__["SelectionDirection"].down });
        }
    };
    return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { "data-automation-type": 'Tablero', tabIndex: tabIndex, className: Object(_TableStyle__WEBPACK_IMPORTED_MODULE_23__["tableWrapperClassName"])(), onFocus: onTableFocus, onBlur: onTableBlur, ref: wrapperRef, onContextMenu: onTableContextMenu, contentEditable: false },
        react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_header_TableHeader__WEBPACK_IMPORTED_MODULE_5__["TableHeader"], { title: title, onTableTitleProposal: onTableTitleProposal, onTableTitleCommit: onMoveToTableHeader, tableTitleId: tableTitleId, tableTitleRef: tableTitleRef, onMoveToTableHeader: onMoveToTableHeader, onMoveSelectionToHost: onMoveSelectionToHost, isReadOnlyMode: isReadOnlyMode }),
        react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { style: { position: 'relative', width: 'fit-content', maxWidth: '100%' } },
            react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { className: Object(_TableStyle__WEBPACK_IMPORTED_MODULE_23__["tableElementWrapperClassName"])(), ref: tableWrapperRef, onScroll: onTableWrapperScroll, onMouseDown: onTableElementWrapperMouseDown },
                Object(_TableAppFeatures__WEBPACK_IMPORTED_MODULE_22__["grabberFeatureEnabled"])() && (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_ColumnGrabbersTable__WEBPACK_IMPORTED_MODULE_20__["ColumnGrabbersTable"], { columns: columns, onInsertColumnProposal: onInsertColumnProposal, onDeleteColumnProposal: onDeleteColumnProposal, getPreset: getPreset, onSortByColumnProposal: onSortByColumnProposal, getLastKnownLocalPresence: getLastKnownLocalPresence, onLocalPresenceValueProposal: onLocalPresenceValueProposal, onCellStyleOverride: onCellStyleOverride, 
                    // Column grabber will only be shown if there is an entry for hovered column Id
                    hoveredTableCellColumnId: hoveredTableCellIds.columnId, tableWrapperElm: tableWrapperRef === null || tableWrapperRef === void 0 ? void 0 : tableWrapperRef.current, selectedColumnId: rowColumnSelectionState.type === _TableDefinitions__WEBPACK_IMPORTED_MODULE_19__["TableElementType"].Column ? rowColumnSelectionState.id : undefined, onRowColumnSelect: onRowColumnSelect, isReadOnlyMode: isReadOnlyMode, activateDragPreview: activateDragPreview, resizeColumnId: resizeColumnProperties.columnId, onResizeColumnLeaveCmd: onResizeColumnLeaveCmd })),
                Object(_TableAppFeatures__WEBPACK_IMPORTED_MODULE_22__["grabberFeatureEnabled"])() && (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_RowGrabbersTable__WEBPACK_IMPORTED_MODULE_21__["RowGrabbersTable"], { rows: rows, rowHeights: rowHeights, onInsertRowProposal: onInsertRowProposal, onDeleteRowProposal: onDeleteRowProposal, onCellStyleOverride: onCellStyleOverride, 
                    // Row grabber will only be shown if there is an entry for hovered row Id
                    hoveredTableCellRowId: hoveredTableCellIds.rowId, getLastKnownLocalPresence: getLastKnownLocalPresence, onLocalPresenceValueProposal: onLocalPresenceValueProposal, selectedRowId: rowColumnSelectionState.type === _TableDefinitions__WEBPACK_IMPORTED_MODULE_19__["TableElementType"].Row ? rowColumnSelectionState.id : undefined, onRowColumnSelect: onRowColumnSelect, isReadOnlyMode: isReadOnlyMode, activateDragPreview: activateDragPreview })),
                react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("table", { tabIndex: tabIndex, className: `${_TableStyle__WEBPACK_IMPORTED_MODULE_23__["tableClassName"]} ${_TableStyle__WEBPACK_IMPORTED_MODULE_23__["tableBorderSpacingClassName"]}`, onKeyDown: onTableKeyDown, ref: tableRef, role: ariaTableRole, "aria-labelledby": tableTitleId, onMouseDown: onTableMouseDown },
                    react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("thead", null,
                        react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("tr", { role: ariaRowRole, ref: headerRowRef }, columns.map((column, index) => {
                            var _a;
                            return (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](react__WEBPACK_IMPORTED_MODULE_0__["Fragment"], { key: column.id },
                                react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_columns_ColumnHeader__WEBPACK_IMPORTED_MODULE_3__["ColumnHeader"], { id: column.id, columnIndex: index, title: column.title, width: Object(_utils_Helper__WEBPACK_IMPORTED_MODULE_12__["getDimensionAsNumber"])(column.width, _StylingConstants__WEBPACK_IMPORTED_MODULE_8__["defaultColumnWidth"]), key: column.id, onColumnSelected: setColumnSelection, 
                                    // Pass through coauthor locations only for header row and current column
                                    coauthorsInColumnHeader: (_a = coauthorsMap.get(_Constants__WEBPACK_IMPORTED_MODULE_11__["headerRowIndex"])) === null || _a === void 0 ? void 0 : _a.get(column.id), onColumnTitleProposal: onColumnTitleProposal, localPresenceLocation: getLocalPresenceLocation(column.id, /*rowId*/ undefined), onColumnResizeProposal: onColumnResizeProposal, onInvokeContextMenu: onInvokeContextMenu, getPreset: getPreset, sendInsertComponentRequestData: sendInsertComponentRequestData, enableUndoRedo: enableUndoRedo, componentContext: componentContext, logger: logger, onMoveSelectionFromCell: onMoveSelectionFromCell, moveSelectionToCellDetails: moveSelectionToCellDetails, cellStyleOverride: Object(_cells_CellHelpers__WEBPACK_IMPORTED_MODULE_18__["getCellStyleOverridesForCell"])(cellStyleOverride, column.id, undefined), onTableCellHover: onTableCellHover, isReadOnlyMode: isReadOnlyMode, onResizeColumnLeaveCmd: onResizeColumnLeaveCmd, 
                                    // If resize column was invoked on a header cell, only the header cell having the matching
                                    // column ID should receive true for resize table column mode.
                                    isResizeColumnMode: !resizeColumnProperties.rowId && resizeColumnProperties.columnId === column.id, tableCellApi: tableCellApi.current, ariaMessage: ariaMessage })));
                        }))),
                    react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("tbody", null, rows.map((row) => (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_rows_TableRow__WEBPACK_IMPORTED_MODULE_4__["TableRow"], { rowId: row.id, data: row.data, columns: columns, key: row.id, onCellValueProposal: onCellValueProposal, rowSpacerColSpan: rowSpacerColSpan, onCellSelectionProposal: onCellSelectionProposal, localPresenceLocation: getLocalPresenceLocation(/*columnId*/ undefined, row.id), 
                        // Pass through coauthor locations only for current row
                        coauthorsInRow: coauthorsMap.get(row.id), onColumnResizeProposal: onColumnResizeProposal, uiContextComponentUrl: `${componentContext.id}`, onInvokeContextMenu: onInvokeContextMenu, componentContext: componentContext, logger: logger, enableUndoRedo: enableUndoRedo, sendInsertComponentRequestData: sendInsertComponentRequestData, onMoveSelectionFromCell: onMoveSelectionFromCell, moveSelectionToCellDetails: moveSelectionToCellDetails, highlightColumnsForClickAttribution: getColumnsForAttribution(row.id, clickHighlightAttributionMap), highlightColumnsForHoverAttribution: getColumnsForAttribution(row.id, hoverHighlightAttributionMap), cellStyleOverride: Object(_cells_CellHelpers__WEBPACK_IMPORTED_MODULE_18__["getCellStyleOverridesForRow"])(cellStyleOverride, row.id), onTableCellHover: onTableCellHover, onNestedComponentDataChangeCallback: onNestedComponentDataChangeCallback, isReadOnlyMode: isReadOnlyMode, onResizeColumnLeaveCmd: onResizeColumnLeaveCmd, 
                        // Only the rows having the matching row ID for resize table column mode should receive
                        // the column Id for resize table column mode.
                        // Rest of the rows receive undefined.
                        resizeColumnId: resizeColumnProperties.rowId === row.id ? resizeColumnProperties.columnId : undefined, updateRowHeights: updateRowHeights, tableCellApi: tableCellApi.current, ariaMessage: ariaMessage }))))),
                isAddRowBtnEnabled && !isReadOnlyMode && (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_rows_AddRow__WEBPACK_IMPORTED_MODULE_13__["AddRow"], { onInsertRow: onInsertRowProposal, rows: rows, label: _StringTable__WEBPACK_IMPORTED_MODULE_14__["addTaskLabel"], newPresenceColumn: columns[taskNameColumnIndex], addRowBtnRef: addRowBtnRef, setFocusAfterAddRowBtn: setFocusAfterAddRowBtn }))),
            enableHorizontalScrollBarGradient.current && (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_TableGradient__WEBPACK_IMPORTED_MODULE_15__["TableGradient"], { gradientSide: _TableGradient__WEBPACK_IMPORTED_MODULE_15__["GradientSide"].Left, width: gradientState.leftWidth, height: gradientState.height, offsetTop: gradientState.offsetTop })),
            enableHorizontalScrollBarGradient.current && (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_TableGradient__WEBPACK_IMPORTED_MODULE_15__["TableGradient"], { gradientSide: _TableGradient__WEBPACK_IMPORTED_MODULE_15__["GradientSide"].Right, width: gradientState.rightWidth, height: gradientState.height, offsetTop: gradientState.offsetTop })),
            Object(_TableAppFeatures__WEBPACK_IMPORTED_MODULE_22__["dragAndDropFeatureEnabled"])() && (react__WEBPACK_IMPORTED_MODULE_0__["createElement"](react__WEBPACK_IMPORTED_MODULE_0__["Suspense"], { fallback: null },
                react__WEBPACK_IMPORTED_MODULE_0__["createElement"](DragManager, { dragAction: dragAction, rows: rows, columns: columns, getTableCellPreview: getTableCellPreview, getRowColumnBounds: getRowColumnBounds })))),
        react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { "aria-live": "polite", "aria-atomic": "true", ref: ariaMessageDivRef, className: _TableStyle__WEBPACK_IMPORTED_MODULE_23__["ariaMessageClassName"] })));
});


/***/ }),

/***/ "../tablero/lib/view/table/Table.props.js":
/*!************************************************!*\
  !*** ../tablero/lib/view/table/Table.props.js ***!
  \************************************************/
/*! exports provided: SearchDirection, SortDirection, DataType, DataTypeMap, RelativeInsertPosition, TableScrollDirection */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SearchDirection", function() { return SearchDirection; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SortDirection", function() { return SortDirection; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "DataType", function() { return DataType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "DataTypeMap", function() { return DataTypeMap; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "RelativeInsertPosition", function() { return RelativeInsertPosition; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "TableScrollDirection", function() { return TableScrollDirection; });
var SearchDirection;
(function (SearchDirection) {
    SearchDirection["Left"] = "left";
    SearchDirection["Right"] = "right";
})(SearchDirection || (SearchDirection = {}));
var SortDirection;
(function (SortDirection) {
    SortDirection["Ascending"] = "ascending";
    SortDirection["Descending"] = "descending";
})(SortDirection || (SortDirection = {}));
var DataType;
(function (DataType) {
    DataType["Text"] = "IStringData";
    DataType["Checkbox"] = "IBooleanData";
    DataType["Date"] = "IDateData";
    DataType["RichText"] = "IRichTextData";
    DataType["SimpleRichText"] = "ISimpleRichText";
})(DataType || (DataType = {}));
/**
 * Maps the column/data type to the corresponding data type field required for nested component creation.
 */
const DataTypeMap = {
    // TODO: Rename data type from rich text to "RichTextTableCell".
    // TODO: May need different data type for different column types - e.g. title/notes.
    [DataType.RichText]: 'RichTextTableCell',
    [DataType.SimpleRichText]: 'SimpleRichTextTableCell',
    [DataType.Date]: 'Date'
};
var RelativeInsertPosition;
(function (RelativeInsertPosition) {
    RelativeInsertPosition["Before"] = "before";
    RelativeInsertPosition["After"] = "after";
})(RelativeInsertPosition || (RelativeInsertPosition = {}));
var TableScrollDirection;
(function (TableScrollDirection) {
    TableScrollDirection["Left"] = "left";
    TableScrollDirection["Right"] = "right";
    TableScrollDirection["Down"] = "down";
    TableScrollDirection["Up"] = "up";
})(TableScrollDirection || (TableScrollDirection = {}));


/***/ }),

/***/ "../tablero/lib/view/table/TableAppFeatures.js":
/*!*****************************************************!*\
  !*** ../tablero/lib/view/table/TableAppFeatures.js ***!
  \*****************************************************/
/*! exports provided: audience, registerSettingsValueTypesAudience, sideMarginFeatureEnabled, grabberFeatureEnabled, dragAndDropFeatureEnabled, resizeColumnFromCmdMenuFeatureEnabled, nestAnyLayoutFeatureEnabled, isRichTextTableEnabled */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "audience", function() { return audience; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "registerSettingsValueTypesAudience", function() { return registerSettingsValueTypesAudience; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "sideMarginFeatureEnabled", function() { return sideMarginFeatureEnabled; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "grabberFeatureEnabled", function() { return grabberFeatureEnabled; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "dragAndDropFeatureEnabled", function() { return dragAndDropFeatureEnabled; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "resizeColumnFromCmdMenuFeatureEnabled", function() { return resizeColumnFromCmdMenuFeatureEnabled; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "nestAnyLayoutFeatureEnabled", function() { return nestAnyLayoutFeatureEnabled; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "isRichTextTableEnabled", function() { return isRichTextTableEnabled; });
/* harmony import */ var _configuration_SettingsValueTypes__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../../configuration/SettingsValueTypes */ "../tablero/lib/configuration/SettingsValueTypes.js");
/* harmony import */ var _fluidx_utilities__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @fluidx/utilities */ "../utilities/lib/settings/settingsValue.js");


let audience;
function registerSettingsValueTypesAudience(hostAudience) {
    audience = hostAudience;
}
// Get settings flag value for side margins
function sideMarginFeatureEnabled() {
    return Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_1__["getSettingsValue"])({
        name: _configuration_SettingsValueTypes__WEBPACK_IMPORTED_MODULE_0__["SettingsValueTypes"].sideMargin,
        defaultValue: true
    });
}
// Get settings flag value for table grabbers
function grabberFeatureEnabled() {
    return Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_1__["getSettingsValue"])({
        name: _configuration_SettingsValueTypes__WEBPACK_IMPORTED_MODULE_0__["SettingsValueTypes"].grabbers,
        defaultValue: true
    });
}
// Get settings flag value for drag and drop feature
function dragAndDropFeatureEnabled() {
    return Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_1__["getSettingsValue"])({
        name: _configuration_SettingsValueTypes__WEBPACK_IMPORTED_MODULE_0__["SettingsValueTypes"].dragAndDrop,
        defaultValue: false
    });
}
// Get settings flag value for resize column from commanding menu (resize column via keyboard).
function resizeColumnFromCmdMenuFeatureEnabled() {
    return (grabberFeatureEnabled() &&
        Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_1__["getSettingsValue"])({
            name: _configuration_SettingsValueTypes__WEBPACK_IMPORTED_MODULE_0__["SettingsValueTypes"].resizeColumnFromCmdMenu,
            defaultValue: Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_1__["enableUpToAudience"])('Production', audience)
        }));
}
function nestAnyLayoutFeatureEnabled() {
    return Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_1__["getSettingsValue"])({
        name: _configuration_SettingsValueTypes__WEBPACK_IMPORTED_MODULE_0__["SettingsValueTypes"].nestAnyLayout,
        defaultValue: Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_1__["enableUpToAudience"])('Staging', audience)
    });
}
function isRichTextTableEnabled() {
    return Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_1__["getSettingsValue"])({
        name: _configuration_SettingsValueTypes__WEBPACK_IMPORTED_MODULE_0__["SettingsValueTypes"].richTextTable,
        defaultValue: Object(_fluidx_utilities__WEBPACK_IMPORTED_MODULE_1__["enableUpToAudience"])('Staging', audience)
    });
}


/***/ }),

/***/ "../tablero/lib/view/table/TableDefinitions.js":
/*!*****************************************************!*\
  !*** ../tablero/lib/view/table/TableDefinitions.js ***!
  \*****************************************************/
/*! exports provided: TableElementType */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "TableElementType", function() { return TableElementType; });
/**
 * Indicates type of selection (row/column) in table.
 */
var TableElementType;
(function (TableElementType) {
    TableElementType[TableElementType["Row"] = 0] = "Row";
    TableElementType[TableElementType["Column"] = 1] = "Column";
})(TableElementType || (TableElementType = {}));


/***/ }),

/***/ "../tablero/lib/view/table/TableGradient.js":
/*!**************************************************!*\
  !*** ../tablero/lib/view/table/TableGradient.js ***!
  \**************************************************/
/*! exports provided: GradientSide, TableGradient */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "GradientSide", function() { return GradientSide; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "TableGradient", function() { return TableGradient; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/react/16.13.1/node_modules/react/index.js");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/styling/7.13.5_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/styling/lib/index.js");


var GradientSide;
(function (GradientSide) {
    GradientSide[GradientSide["Left"] = 0] = "Left";
    GradientSide[GradientSide["Right"] = 1] = "Right";
})(GradientSide || (GradientSide = {}));
const gradientClassName = (gradientSide, width, height, offsetTop) => {
    return Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["mergeStyles"])({
        // zIndex is required to have stack order greater than or equal to that of Table cells as gradient should always be visible above the cells.
        zIndex: 1,
        position: 'absolute',
        pointerEvents: 'none',
        display: 'inline',
        height: `${height}px`,
        top: `${offsetTop}px`,
        width: `${width}px`,
        left: gradientSide === GradientSide.Left ? '0px' : 'auto',
        right: gradientSide === GradientSide.Right ? '0px' : 'auto',
        backgroundImage: `linear-gradient(to ${gradientSide === GradientSide.Left ? 'right' : 'left'}, #EAEAEA, transparent)`
    });
};
const TableGradient = react__WEBPACK_IMPORTED_MODULE_0__["memo"](function TableGradient(props) {
    return react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", { className: gradientClassName(props.gradientSide, props.width, props.height, props.offsetTop) });
});


/***/ }),

/***/ "../tablero/lib/view/table/TableKeyboarding.js":
/*!*****************************************************!*\
  !*** ../tablero/lib/view/table/TableKeyboarding.js ***!
  \*****************************************************/
/*! exports provided: getKeyCodeFromSelectionDirections, getTableKeyboardHandler */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getKeyCodeFromSelectionDirections", function() { return getKeyCodeFromSelectionDirections; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getTableKeyboardHandler", function() { return getTableKeyboardHandler; });
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/utilities/7.22.0_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/utilities/lib/KeyCodes.js");
/* harmony import */ var _Table_props__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./Table.props */ "../tablero/lib/view/table/Table.props.js");
/* harmony import */ var _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @fluidx/office-fluid-types */ "../office-fluid-types/lib/ComponentContracts/ComponentSelection.js");
/* harmony import */ var _ms_office_bohemia_ux__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @ms/office-bohemia-ux */ "../office-bohemia-ux/lib/eventing/matchEventWithOrigin.js");




const keyCodesToPreventDefault = new Set([
    // Prevent scroll left, required for read only mode.
    office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["KeyCodes"].left,
    // Prevent scroll right, required for read only mode.
    office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["KeyCodes"].right,
    // Prevent scroll down
    office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["KeyCodes"].down,
    // Prevent scroll up
    office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["KeyCodes"].up,
    // Prevent the browser moving focus
    office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["KeyCodes"].tab
]);
const getKeyCodeFromSelectionDirections = {};
getKeyCodeFromSelectionDirections[_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_2__["SelectionDirection"].left] = office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["KeyCodes"].left;
getKeyCodeFromSelectionDirections[_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_2__["SelectionDirection"].right] = office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["KeyCodes"].right;
getKeyCodeFromSelectionDirections[_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_2__["SelectionDirection"].up] = office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["KeyCodes"].up;
getKeyCodeFromSelectionDirections[_fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_2__["SelectionDirection"].down] = office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["KeyCodes"].down;
const lastElementMatchesId = (arr, id) => {
    if (arr.length > 0) {
        if (arr[arr.length - 1].id === id) {
            return true;
        }
    }
    return false;
};
const moveSelectionLeft = (localPresenceLocation, rows, columns, setLocalPresenceDetails, allowWrapping, focusDispatcher, target, updateTableScroll, selectionParams) => {
    const currentColIndex = columns.findIndex((c) => c.id === localPresenceLocation.columnId);
    let newColIndex = currentColIndex;
    if (currentColIndex < 0) {
        return;
    }
    // if we're on the very first cell in a table and we want to move left, we should give focus
    // back to the hosting canvas and exit out of this method.
    // The table's logic for setting cell selection to undefined when blurred is what clears the
    // cell presence upon this dispatch happening.
    if (currentColIndex === 0 && localPresenceLocation.type === 'column') {
        if (focusDispatcher({ direction: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_2__["SelectionDirection"].left })) {
            return;
        }
    }
    if (currentColIndex > 0) {
        // The simple case is that we aren't in the first column of the row, we just move selection to the previous column.
        setLocalPresenceDetails(Object.assign(Object.assign({}, localPresenceLocation), { columnId: columns[currentColIndex - 1].id }), selectionParams);
        newColIndex = currentColIndex - 1;
    }
    else if (currentColIndex === 0 && localPresenceLocation.type === 'cell' && allowWrapping) {
        // If we were on the first column and we want to allow wrapping, move the selection to the last column of the previous row.
        const currentRowIndex = rows.findIndex((r) => r.id === localPresenceLocation.rowId);
        if (currentRowIndex > 0) {
            // If there is a previous row, move selection to the last column of the previous row.
            setLocalPresenceDetails(Object.assign(Object.assign({}, localPresenceLocation), { columnId: columns[columns.length - 1].id, rowId: rows[currentRowIndex - 1].id }), selectionParams);
        }
        else if (currentRowIndex === 0) {
            // If we are on the first row, move the selection to the last column header.
            setLocalPresenceDetails({
                type: 'column',
                columnId: columns[columns.length - 1].id
            }, selectionParams);
        }
        newColIndex = columns.length - 1;
    }
    // If the cell we navigated to is in a different column then move table scrollbar if needed
    if (newColIndex !== currentColIndex && target !== undefined && updateTableScroll !== undefined) {
        updateTableScroll(target, newColIndex, _Table_props__WEBPACK_IMPORTED_MODULE_1__["TableScrollDirection"].Left);
    }
};
const moveSelectionRight = (localPresenceLocation, rows, columns, setLocalPresenceDetails, allowWrapping, focusDispatcher, target, updateTableScroll, onInsertRowProposal, selectionParams, focusOnAddRowButton) => {
    const currentColIndex = columns.findIndex((c) => c.id === localPresenceLocation.columnId);
    let newColIndex = currentColIndex;
    if (currentColIndex < 0) {
        return;
    }
    if (currentColIndex === columns.length - 1 && localPresenceLocation.type === 'cell') {
        if (lastElementMatchesId(rows, localPresenceLocation.rowId)) {
            // if we're on the very last cell in the table, and pressed tab, we want to insert a new row and
            // move focus to the first cell of the new row. We only want to do so if indicated by having the insert row function defined.
            if (onInsertRowProposal !== undefined) {
                const insertArgument = {
                    type: 'insertRelative',
                    id: localPresenceLocation.rowId,
                    insertPosition: _Table_props__WEBPACK_IMPORTED_MODULE_1__["RelativeInsertPosition"].After,
                    newPresenceLocation: columns[0]
                };
                onInsertRowProposal(insertArgument);
            }
            else {
                // if we're in the very last cell in the table, and did not press tab, we should give focus
                // back to the hosting canvas and exit out of this method.
                // The table's logic for setting cell selection to undefined when blurred is what clears the
                // cell presence upon this dispatch happening.
                // Set focus to 'Add Row Button' if it's enabled.
                if (focusOnAddRowButton) {
                    focusOnAddRowButton();
                    return;
                }
                if (focusDispatcher({ direction: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_2__["SelectionDirection"].right })) {
                    return;
                }
            }
        }
    }
    if (currentColIndex + 1 < columns.length) {
        // The simple case is that we aren't in the last column of the row, then just move the selection to the next row.
        setLocalPresenceDetails(Object.assign(Object.assign({}, localPresenceLocation), { columnId: columns[currentColIndex + 1].id }), selectionParams);
        newColIndex = currentColIndex + 1;
    }
    else if (currentColIndex === columns.length - 1 && allowWrapping) {
        // If we are in the last column of the row and we want to allow wrapping, move selection to the first column of the next row.
        if (localPresenceLocation.type === 'column') {
            if (rows.length > 0) {
                // If the selection is currently on the last tab header, move it to the first data cell.
                setLocalPresenceDetails({
                    type: 'cell',
                    rowId: rows[0].id,
                    columnId: columns[0].id
                }, selectionParams);
            }
            else if (rows.length === 0 && focusOnAddRowButton) {
                focusOnAddRowButton();
            }
        }
        else if (localPresenceLocation.type === 'cell') {
            // If the selection is on a cell that is last column of a row that isn't the last row, move it to the first column of the next row
            const currentRowIndex = rows.findIndex((r) => r.id === localPresenceLocation.rowId);
            if (currentRowIndex < 0) {
                return;
            }
            if (currentRowIndex + 1 < rows.length) {
                setLocalPresenceDetails(Object.assign(Object.assign({}, localPresenceLocation), { columnId: columns[0].id, rowId: rows[currentRowIndex + 1].id }), selectionParams);
            }
        }
        newColIndex = 0;
    }
    // If the cell we navigated to is in a different column then move table scrollbar if needed
    if (newColIndex !== currentColIndex && target && updateTableScroll) {
        updateTableScroll(target, newColIndex, _Table_props__WEBPACK_IMPORTED_MODULE_1__["TableScrollDirection"].Right);
    }
};
const moveSelectionDown = (localPresenceLocation, rows, setLocalPresenceDetails, focusDispatcher, getCaretHorizontalPosition, selectionParams, focusOnAddRowButton) => {
    // If a column is selected, move selection down to the cell in the same column of the first row.
    if (localPresenceLocation.type === 'column') {
        if (rows.length > 0) {
            setLocalPresenceDetails({
                type: 'cell',
                rowId: rows[0].id,
                columnId: localPresenceLocation.columnId
            }, selectionParams);
        }
        else if (rows.length === 0 && focusOnAddRowButton) {
            focusOnAddRowButton();
        }
    }
    else if (localPresenceLocation.type === 'cell') {
        // Try to move selection down to the next row.
        const currentRowIndex = rows.findIndex((r) => r.id === localPresenceLocation.rowId);
        if (currentRowIndex < 0) {
            return;
        }
        if (currentRowIndex + 1 < rows.length) {
            setLocalPresenceDetails({
                type: 'cell',
                rowId: rows[currentRowIndex + 1].id,
                columnId: localPresenceLocation.columnId
            }, selectionParams);
        }
        else if (currentRowIndex + 1 === rows.length) {
            // If row index + 1 is equal to rows length that means we pressed keydown on the last arros
            // we should try to move out of tablero and give focus to hosting canvas
            // Set focus to 'Add Row Button' if it's enabled.
            if (focusOnAddRowButton) {
                focusOnAddRowButton();
                return;
            }
            // This below code dispatches focus if we have scriptor below Table/Task otherwise it returns false, like in OWA.
            if (focusDispatcher({
                direction: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_2__["SelectionDirection"].down,
                currentCoordinates: {
                    x: getCaretHorizontalPosition()
                }
            })) {
                return;
            }
        }
    }
};
const moveSelectionUp = (localPresenceLocation, rows, setLocalPresenceDetails, onMoveToTableTitle, selectionParams) => {
    if (localPresenceLocation.type === 'cell') {
        const currentRowIndex = rows.findIndex((r) => r.id === localPresenceLocation.rowId);
        if (currentRowIndex > 0) {
            // Try to move selection up to the previous row.
            setLocalPresenceDetails({
                type: 'cell',
                columnId: localPresenceLocation.columnId,
                rowId: rows[currentRowIndex - 1].id
            }, selectionParams);
        }
        else if (currentRowIndex === 0) {
            // If the currently selected cell is in the first row, move selection to the column header.
            setLocalPresenceDetails({
                type: 'column',
                columnId: localPresenceLocation.columnId
            }, selectionParams);
        }
    }
    else if (localPresenceLocation.type === 'column') {
        // If we arrow up in a table header move focus to the tablet title
        onMoveToTableTitle();
    }
};
const getTableKeyboardHandler = (localPresenceLocation, setLocalPresenceDetails, rows, columns, focusDispatcher, updateTableScroll, onInsertRowProposal, getCaretHorizontalPosition, onMoveToTableTitle, onUndo, onRedo, focusOnAddRowButton) => (ev) => {
    if (!localPresenceLocation) {
        return;
    }
    if (!ev) {
        throw new Error('getTableKeyboardHandler returns a function that expects either an event.');
    }
    let keyCode;
    if (ev && ev.nativeEvent && ev.currentTarget) {
        // TODO:TABLERO: wrap this check in a utility function if we end up using it in other places
        // If the composedPath property is not found on the event (some browsers do not support) build the path by traversing parent list
        const composedPath = ev.nativeEvent.composedPath ? ev.nativeEvent.composedPath() : Object(_ms_office_bohemia_ux__WEBPACK_IMPORTED_MODULE_3__["buildEventPath"])(ev.nativeEvent);
        if (Object(_ms_office_bohemia_ux__WEBPACK_IMPORTED_MODULE_3__["isEventFromComponentHostingElement"])(ev.currentTarget, composedPath)) {
            return;
        }
        if (keyCodesToPreventDefault.has(ev.keyCode)) {
            if (ev.preventDefault) {
                ev.preventDefault();
            }
        }
    }
    // Set Key Code
    keyCode = ev.keyCode;
    switch (keyCode) {
        case office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["KeyCodes"].left: {
            moveSelectionLeft(localPresenceLocation, rows, columns, setLocalPresenceDetails, true, focusDispatcher, ev === null || ev === void 0 ? void 0 : ev.target, updateTableScroll, { mode: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_2__["SelectionMode"].ip, direction: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_2__["SelectionDirection"].left });
            break;
        }
        case office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["KeyCodes"].right: {
            moveSelectionRight(localPresenceLocation, rows, columns, setLocalPresenceDetails, true, focusDispatcher, ev === null || ev === void 0 ? void 0 : ev.target, updateTableScroll, undefined, { mode: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_2__["SelectionMode"].ip, direction: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_2__["SelectionDirection"].right }, focusOnAddRowButton);
            break;
        }
        case office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["KeyCodes"].down: {
            moveSelectionDown(localPresenceLocation, rows, setLocalPresenceDetails, focusDispatcher, getCaretHorizontalPosition, { mode: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_2__["SelectionMode"].ip, direction: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_2__["SelectionDirection"].down }, focusOnAddRowButton);
            break;
        }
        case office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["KeyCodes"].enter: {
            moveSelectionDown(localPresenceLocation, rows, setLocalPresenceDetails, focusDispatcher, getCaretHorizontalPosition, { mode: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_2__["SelectionMode"].ip, direction: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_2__["SelectionDirection"].down }, focusOnAddRowButton);
            break;
        }
        case office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["KeyCodes"].up: {
            moveSelectionUp(localPresenceLocation, rows, setLocalPresenceDetails, onMoveToTableTitle, {
                mode: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_2__["SelectionMode"].ip,
                direction: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_2__["SelectionDirection"].up
            });
            break;
        }
        case office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["KeyCodes"].tab: {
            if (ev === null || ev === void 0 ? void 0 : ev.shiftKey) {
                moveSelectionLeft(localPresenceLocation, rows, columns, setLocalPresenceDetails, true, focusDispatcher, ev === null || ev === void 0 ? void 0 : ev.target, updateTableScroll, {
                    mode: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_2__["SelectionMode"].ip,
                    direction: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_2__["SelectionDirection"].left
                });
            }
            else {
                moveSelectionRight(localPresenceLocation, rows, columns, setLocalPresenceDetails, true, focusDispatcher, ev === null || ev === void 0 ? void 0 : ev.target, updateTableScroll, onInsertRowProposal, {
                    mode: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_2__["SelectionMode"].ip,
                    direction: _fluidx_office_fluid_types__WEBPACK_IMPORTED_MODULE_2__["SelectionDirection"].right
                });
            }
            break;
        }
        case office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["KeyCodes"].z: {
            if (ev === null || ev === void 0 ? void 0 : ev.ctrlKey) {
                onUndo && onUndo();
            }
            break;
        }
        case office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["KeyCodes"].y: {
            if (ev === null || ev === void 0 ? void 0 : ev.ctrlKey) {
                onRedo && onRedo();
            }
            break;
        }
    }
};


/***/ }),

/***/ "../tablero/lib/view/table/TableStyle.js":
/*!***********************************************!*\
  !*** ../tablero/lib/view/table/TableStyle.js ***!
  \***********************************************/
/*! exports provided: tableWrapperClassName, tableBorderSpacingClassName, tableClassName, tableElementWrapperClassName, ariaMessageClassName */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableWrapperClassName", function() { return tableWrapperClassName; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableBorderSpacingClassName", function() { return tableBorderSpacingClassName; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableClassName", function() { return tableClassName; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableElementWrapperClassName", function() { return tableElementWrapperClassName; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ariaMessageClassName", function() { return ariaMessageClassName; });
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/styling/7.13.5_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/styling/lib/index.js");
/* harmony import */ var _StylingConstants__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../StylingConstants */ "../tablero/lib/view/StylingConstants.js");
/* harmony import */ var _TableAppFeatures__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./TableAppFeatures */ "../tablero/lib/view/table/TableAppFeatures.js");



let tableWrapperClassNameValue;
let tableElementWrapperClassNameValue;
function tableWrapperClassName() {
    if (tableWrapperClassNameValue === undefined) {
        tableWrapperClassNameValue = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["mergeStyles"])({
            verticalAlign: 'bottom',
            display: 'inline-block',
            selectors: {
                ':focus': {
                    outline: 'none'
                }
            },
            maxWidth: Object(_TableAppFeatures__WEBPACK_IMPORTED_MODULE_2__["sideMarginFeatureEnabled"])() ? `calc(100% - ${_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["tableroLeftMargin"] + _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["tableroRightMargin"]}px)` : '100%',
            backgroundColor: _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["colorLabels"].tableWrapperBackgroundColor,
            marginLeft: Object(_TableAppFeatures__WEBPACK_IMPORTED_MODULE_2__["sideMarginFeatureEnabled"])() ? _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["tableroLeftMargin"] : 0,
            marginRight: Object(_TableAppFeatures__WEBPACK_IMPORTED_MODULE_2__["sideMarginFeatureEnabled"])() ? _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["tableroRightMargin"] : 0,
            marginBottom: Object(_TableAppFeatures__WEBPACK_IMPORTED_MODULE_2__["grabberFeatureEnabled"])() ? _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["tableroBottomMargin"] : 0,
            // Absolute positioned elements (TableGradient in this case) use
            // nearest ancestor with non-static position as their reference.
            position: 'relative'
        });
    }
    return tableWrapperClassNameValue;
}
const tableBorderSpacingClassName = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["mergeStyles"])({
    borderSpacing: `${_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperBorderWidth"]}px`,
    selectors: {
        // Fix for - Bug 4218954: Table borders are missing at 85% zoom.
        // For browsers, if (device-pixel-ratio * borderSpacing) < 1, then it is rounded off to zero.
        // So borderSpacing is adjusted when device-pixel-ratio changes, to keep it greater than 1, less than 2.
        '@media (max-resolution: 95dpi)': {
            // 96 dpi - 1
            borderSpacing: `${_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperBorderWidthZoomLessThanOne"]}px`
        },
        '@media (max-resolution: 47dpi)': {
            // 96 dpi / 2 - 1
            borderSpacing: `${_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperBorderWidthZoomLessThanHalf"]}px`
        },
        '@media (max-resolution: 31dpi)': {
            // 96 dpi / 3 - 1
            borderSpacing: `${_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperBorderWidthZoomLessThanThird"]}px`
        }
    }
});
const tableClassName = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["mergeStyles"])({
    tableLayout: 'fixed',
    // Browsers don't respect fixed table layout if width is not specified.
    // width: 'fit-content' - chrome respect, firefox and spartan fail
    // width: 'max-content', chrome and firefox respect, spartan fail
    // width: 'unset'/'initial'/'100%', chrome fail
    // This leaves giving absolute width as only option.
    // width: <calculated width of table> - horizontal scroll bar of ~2 px.
    // Using '0px' works fine on all browsers. Width of table
    // will auto inflate later when kids of <table> are rendered.
    width: '0px',
    // Using '0' will help auto inflate later when child cell components are rendered.
    height: '0',
    borderCollapse: 'separate !important',
    backgroundColor: _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["colorLabels"].tableBackgroundColor,
    borderRadius: `${_StylingConstants__WEBPACK_IMPORTED_MODULE_1__["tableBorderRadius"]}px`,
    selectors: {
        ':focus': {
            outline: 'none'
        },
        [office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["HighContrastSelector"]]: {
            backgroundColor: 'windowText',
            MsHighContrastAdjust: 'none'
        }
    },
    marginBottom: _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["tableBottomMargin"],
    marginRight: _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["tableRightMargin"]
});
function tableElementWrapperClassName() {
    if (tableElementWrapperClassNameValue === undefined) {
        tableElementWrapperClassNameValue = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["mergeStyles"])({
            // If grabber feature is disabled, set the top padding to "grabber height + top border height"
            // If grabber feature is enabled, don't put top padding as grabber will take that entire area
            paddingTop: Object(_TableAppFeatures__WEBPACK_IMPORTED_MODULE_2__["grabberFeatureEnabled"])() ? '' : _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["columnGrabberHeight"] + _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["cellWrapperBorderWidth"],
            maxWidth: '100%',
            overflowX: 'auto',
            overflowY: 'hidden',
            selectors: {
                '::-webkit-scrollbar': {
                    height: _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["tableScrollBarHeight"],
                    width: _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["tableScrollBarWidth"]
                },
                '::-webkit-scrollbar-thumb': {
                    background: _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["tableScrollBarColor"],
                    borderRadius: _StylingConstants__WEBPACK_IMPORTED_MODULE_1__["tableScrollBarBorderRadius"]
                }
            }
        });
    }
    return tableElementWrapperClassNameValue;
}
const ariaMessageClassName = Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_0__["mergeStyles"])({
    height: '0px',
    width: '0px',
    overflow: 'hidden'
});


/***/ }),

/***/ "../tablero/lib/view/utils/Helper.js":
/*!*******************************************!*\
  !*** ../tablero/lib/view/utils/Helper.js ***!
  \*******************************************/
/*! exports provided: getDimensionAsNumber, isCaretAtStart, isCaretAtEnd, getEdgeVisibilityOfElementInContainer, isElementCompletelyOutsideContainer, getResizeColumnCalloutBounds */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getDimensionAsNumber", function() { return getDimensionAsNumber; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "isCaretAtStart", function() { return isCaretAtStart; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "isCaretAtEnd", function() { return isCaretAtEnd; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getEdgeVisibilityOfElementInContainer", function() { return getEdgeVisibilityOfElementInContainer; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "isElementCompletelyOutsideContainer", function() { return isElementCompletelyOutsideContainer; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getResizeColumnCalloutBounds", function() { return getResizeColumnCalloutBounds; });
/* harmony import */ var _StylingConstants__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../StylingConstants */ "../tablero/lib/view/StylingConstants.js");

// This function is being currently used to ensure that table width is a number
// Due to a commit, the default width was set to `157px` instead of `157`.
// This resulted in calculation errors during resize, and the resultant width being a invalid string
// Fix has been done for new table, and this code is required only for the tables affected between the two commits.
// TODO:This can be removed once the min-width functionality is introduced
function getDimensionAsNumber(value, defaultValue) {
    return isNaN(value) ? defaultValue : value;
}
// Takes in an input and checks whether the caret position is at the beginning.
const isCaretAtStart = (inputEl) => {
    if (inputEl) {
        return (inputEl === null || inputEl === void 0 ? void 0 : inputEl.selectionStart) === 0 && (inputEl === null || inputEl === void 0 ? void 0 : inputEl.selectionEnd) === 0;
    }
    return false;
};
// Takes in an input and checks whether the caret position is at the end.
const isCaretAtEnd = (inputEl) => {
    if (inputEl) {
        const inputElLength = inputEl === null || inputEl === void 0 ? void 0 : inputEl.value.length;
        return (inputEl === null || inputEl === void 0 ? void 0 : inputEl.selectionStart) === inputElLength && (inputEl === null || inputEl === void 0 ? void 0 : inputEl.selectionEnd) === inputElLength;
    }
    return false;
};
/**
 * Get visibility of given element's left and right edges with reference to the container element
 */
const getEdgeVisibilityOfElementInContainer = (inputElm, containerElm) => {
    const inputElmBounds = inputElm === null || inputElm === void 0 ? void 0 : inputElm.getBoundingClientRect();
    const containerBounds = containerElm === null || containerElm === void 0 ? void 0 : containerElm.getBoundingClientRect();
    let isLeftEdgeVisible = false;
    let isRightEdgeVisible = false;
    if (inputElmBounds && containerBounds) {
        isLeftEdgeVisible = containerBounds.left <= inputElmBounds.left && inputElmBounds.left <= containerBounds.right;
        isRightEdgeVisible = containerBounds.left <= inputElmBounds.right && inputElmBounds.right <= containerBounds.right;
    }
    return {
        isLeftEdgeVisible,
        isRightEdgeVisible
    };
};
/**
 * Checks if given element is completely outside viewport with reference to the container element
 */
const isElementCompletelyOutsideContainer = (inputElm, containerElm) => {
    const inputElmBounds = inputElm === null || inputElm === void 0 ? void 0 : inputElm.getBoundingClientRect();
    const containerBounds = containerElm === null || containerElm === void 0 ? void 0 : containerElm.getBoundingClientRect();
    // If we don't have both the bounds, we can't calculate visibility
    if (inputElmBounds && containerBounds) {
        if (inputElmBounds.x + inputElmBounds.width < containerBounds.x ||
            inputElmBounds.x > containerBounds.x + containerBounds.width) {
            return true;
        }
    }
    return false;
};
/**
 * Helper function to get resize column callout beak's diagonal value.
 */
const getResizeColumnCalloutDiagonal = () => {
    let beakRectHeight = _StylingConstants__WEBPACK_IMPORTED_MODULE_0__["resizeColumnCalloutStyleConstants"].beakWidth;
    let beakRectWidth = _StylingConstants__WEBPACK_IMPORTED_MODULE_0__["resizeColumnCalloutStyleConstants"].beakWidth;
    return Math.sqrt(Math.pow(beakRectHeight, 2) + Math.pow(beakRectWidth, 2));
};
/**
 * Helper function to calculate correct bounds for resize column callout based on TableWrapperElm client rect.
 */
const getResizeColumnCalloutBounds = (tableWrapperClientRect) => {
    // Fixing callout top bound as callout is placed higher than the table wrapper.
    return {
        left: tableWrapperClientRect.left,
        top: tableWrapperClientRect.top -
            _StylingConstants__WEBPACK_IMPORTED_MODULE_0__["resizeColumnCalloutStyleConstants"].gapSpace -
            getResizeColumnCalloutDiagonal() / 2 -
            _StylingConstants__WEBPACK_IMPORTED_MODULE_0__["resizeColumnCalloutStyleConstants"].height,
        width: tableWrapperClientRect.width,
        height: tableWrapperClientRect.height,
        right: tableWrapperClientRect.right,
        bottom: tableWrapperClientRect.bottom
    };
};


/***/ }),

/***/ "../tablero/lib/view/utils/ResizeObserverWrapper.js":
/*!**********************************************************!*\
  !*** ../tablero/lib/view/utils/ResizeObserverWrapper.js ***!
  \**********************************************************/
/*! exports provided: ensureResizeObserver, default */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ensureResizeObserver", function() { return ensureResizeObserver; });
// Call this function to add polyfill for ResizeObserver in unsupported browsers
async function ensureResizeObserver() {
    if (typeof ResizeObserver === 'undefined') {
        const module = await __webpack_require__.e(/*! import() | ResizeObserverPolyfill */ "ResizeObserverPolyfill").then(__webpack_require__.bind(null, /*! resize-observer-polyfill */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/resize-observer-polyfill/1.5.1/node_modules/resize-observer-polyfill/dist/ResizeObserver.es.js"));
        window.ResizeObserver = module.default;
    }
}
// Make a call to ensureResizeObserver before using this API
// e.g. ensureResizeObserver().then(() => { // Code using ResizeObserverWrapper() })
class ResizeObserverWrapper {
    constructor(callback) {
        this.resizeObserver = undefined;
        try {
            this.resizeObserver = new ResizeObserver(callback);
        }
        catch (_a) {
            throw 'Error in ResizeObserver';
        }
    }
    observe(target) {
        var _a;
        (_a = this.resizeObserver) === null || _a === void 0 ? void 0 : _a.observe(target);
    }
    unobserve(target) {
        var _a;
        (_a = this.resizeObserver) === null || _a === void 0 ? void 0 : _a.unobserve(target);
    }
    disconnect() {
        var _a;
        (_a = this.resizeObserver) === null || _a === void 0 ? void 0 : _a.disconnect();
    }
}
/* harmony default export */ __webpack_exports__["default"] = (ResizeObserverWrapper);


/***/ }),

/***/ "../tablero/lib/view/utils/TableCellHighContrastUtility.js":
/*!*****************************************************************!*\
  !*** ../tablero/lib/view/utils/TableCellHighContrastUtility.js ***!
  \*****************************************************************/
/*! exports provided: getHCBoxShadowForGrabberSelection, getCellEdgeStyleForHighContrast */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getHCBoxShadowForGrabberSelection", function() { return getHCBoxShadowForGrabberSelection; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getCellEdgeStyleForHighContrast", function() { return getCellEdgeStyleForHighContrast; });
/* harmony import */ var _StylingConstants__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../StylingConstants */ "../tablero/lib/view/StylingConstants.js");
/* harmony import */ var office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! office-ui-fabric-react */ "../../common/temp/node_modules/.pnpm/office.pkgs.visualstudio.com/@uifabric/styling/7.13.5_c8cfd8d2632af57a582ce3bb7e56b488/node_modules/@uifabric/styling/lib/index.js");
/* harmony import */ var _zIndex__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../zIndex */ "../tablero/lib/view/zIndex.js");



/**
 * Applies Highlighted box shadow on Top of each cell for allEdges of Column and to the Right for allEdges for Row
 * @param isCellSelected
 * @param cellStyleOverride
 */
const getHCBoxShadowForGrabberSelection = (isCellSelected, cellStyleOverride) => {
    var _a;
    return !isCellSelected && (cellStyleOverride === null || cellStyleOverride === void 0 ? void 0 : cellStyleOverride.allEdges)
        ? ((_a = cellStyleOverride === null || cellStyleOverride === void 0 ? void 0 : cellStyleOverride.allEdges) === null || _a === void 0 ? void 0 : _a.columnId) ? _StylingConstants__WEBPACK_IMPORTED_MODULE_0__["columnHCBoxShadow"]
            : _StylingConstants__WEBPACK_IMPORTED_MODULE_0__["rowHCBoxShadow"]
        : 'none';
};
/**
 * compares the given Edge type with the edge layout properties like top, bottom, left, right
 * If matches, set it to 0, otherwise unset.
 * @param givenEdgeType
 * @param edgeLayout
 */
const setEdgePositionLayout = (givenEdgeType, edgeLayout) => {
    return givenEdgeType === edgeLayout ? 0 : 'unset';
};
/**
 * returns the exact cell height (including borders, margin, padding)
 * @param tableCellWrapperRef
 */
const getHeaderCellHeight = (tableCellWrapperRef) => {
    var _a;
    return ((_a = tableCellWrapperRef === null || tableCellWrapperRef === void 0 ? void 0 : tableCellWrapperRef.current) === null || _a === void 0 ? void 0 : _a.getBoundingClientRect().height) + 'px';
};
/**
 * Sets top, bottom, left, right properties based upon which edge to display
 * This is done to give outline effect to selective edges.
 * @param type
 * @param tableCellWrapperRef
 * @param isHeaderCell
 */
const getCellEdgeStyleForHighContrast = (type, tableCellWrapperRef, isHeaderCell) => {
    return Object(office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["mergeStyles"])({
        selectors: {
            [office_ui_fabric_react__WEBPACK_IMPORTED_MODULE_1__["HighContrastSelector"]]: {
                display: 'block',
                width: type === 0 /* Left */ || type === 1 /* Right */ ? '0' : '100%',
                height: 
                // getHeaderCellHeight is used for HeaderCell as it contains border bottom of 1px and 100% height doesn't include that border height.
                type === 0 /* Left */ || type === 1 /* Right */
                    ? `${isHeaderCell ? getHeaderCellHeight(tableCellWrapperRef) : '100%'}`
                    : '0',
                position: 'absolute',
                // set right, if edge type is right.
                right: setEdgePositionLayout(type, 1 /* Right */),
                // set left, if edge type is left.
                left: setEdgePositionLayout(type, 0 /* Left */),
                // set top, if edge type is top.
                top: setEdgePositionLayout(type, 2 /* Top */),
                // set bottom, if edge type is bottom.
                bottom: setEdgePositionLayout(type, 3 /* Bottom */),
                zIndex: _zIndex__WEBPACK_IMPORTED_MODULE_2__["selectedCellZIndex"],
                outline: _StylingConstants__WEBPACK_IMPORTED_MODULE_0__["highContrastOutline"]
            }
        }
    });
};


/***/ }),

/***/ "../tablero/lib/view/utils/getBoxShadow.js":
/*!*************************************************!*\
  !*** ../tablero/lib/view/utils/getBoxShadow.js ***!
  \*************************************************/
/*! exports provided: boxShadowCellType, getBoxShadow */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "boxShadowCellType", function() { return boxShadowCellType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getBoxShadow", function() { return getBoxShadow; });
/* harmony import */ var _StylingConstants__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../StylingConstants */ "../tablero/lib/view/StylingConstants.js");

var boxShadowCellType;
(function (boxShadowCellType) {
    boxShadowCellType[boxShadowCellType["columnGrabber"] = 0] = "columnGrabber";
    boxShadowCellType[boxShadowCellType["rowGrabber"] = 1] = "rowGrabber";
    boxShadowCellType[boxShadowCellType["columnAllEdges"] = 2] = "columnAllEdges";
    boxShadowCellType[boxShadowCellType["rowBottomEdge"] = 3] = "rowBottomEdge";
    boxShadowCellType[boxShadowCellType["rowLastEdge"] = 4] = "rowLastEdge";
})(boxShadowCellType || (boxShadowCellType = {}));
/**
 * return box shadow string according to boxShadowCellType, color and width arguments
 * @param args
 */
const getBoxShadow = (args) => {
    const defaultBoxShadow = `0 0 0 ${args.width}px ${args.color}`;
    switch (args.type) {
        case boxShadowCellType.columnGrabber:
        case boxShadowCellType.columnAllEdges:
            return `${defaultBoxShadow}, ${_StylingConstants__WEBPACK_IMPORTED_MODULE_0__["selectedCellBoxShadow"]}`;
        case boxShadowCellType.rowGrabber:
            return `${defaultBoxShadow}, ${_StylingConstants__WEBPACK_IMPORTED_MODULE_0__["rowBottomBoxShadow"]}, ${_StylingConstants__WEBPACK_IMPORTED_MODULE_0__["rowLeftBoxShadow"]}`;
        case boxShadowCellType.rowBottomEdge:
            return `${defaultBoxShadow}, ${_StylingConstants__WEBPACK_IMPORTED_MODULE_0__["rowBottomBoxShadow"]}`;
        case boxShadowCellType.rowLastEdge:
            return `${defaultBoxShadow}, ${_StylingConstants__WEBPACK_IMPORTED_MODULE_0__["rowBottomBoxShadow"]}, ${_StylingConstants__WEBPACK_IMPORTED_MODULE_0__["rowRightBoxShadow"]}`;
    }
};


/***/ }),

/***/ "../tablero/lib/view/utils/isHighContrastModeOn.js":
/*!*********************************************************!*\
  !*** ../tablero/lib/view/utils/isHighContrastModeOn.js ***!
  \*********************************************************/
/*! exports provided: isHighContrastModeOn, highContrastMediaObserver */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "isHighContrastModeOn", function() { return isHighContrastModeOn; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "highContrastMediaObserver", function() { return highContrastMediaObserver; });
let isHighContrastModeOn = checkHighContrastMode();
// Check if Windows's high contrast mode is on.
// Works in Edge.
function checkHighContrastMode() {
    const queryList = window.matchMedia && window.matchMedia('(-ms-high-contrast: active)');
    return (queryList === null || queryList === void 0 ? void 0 : queryList.matches) || false;
}
const highContrastMediaObserver = window.matchMedia && window.matchMedia('(-ms-high-contrast: active)');
highContrastMediaObserver.addListener(() => (isHighContrastModeOn = checkHighContrastMode()));


/***/ }),

/***/ "../tablero/lib/view/zIndex.js":
/*!*************************************!*\
  !*** ../tablero/lib/view/zIndex.js ***!
  \*************************************/
/*! exports provided: addRowColumnButtonZIndex, columnResizeHandleZIndex, tableTitleInvisibleTextZIndex, tableTitleUnderlineOverlayZIndex, tableTitleInputElementZIndex, suggestionsMenuZIndex, personaCardZIndex, personaBubbleZIndex, dragPreviewZIndex, tableGrabberMenuZIndex, selectedCellZIndex */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "addRowColumnButtonZIndex", function() { return addRowColumnButtonZIndex; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "columnResizeHandleZIndex", function() { return columnResizeHandleZIndex; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableTitleInvisibleTextZIndex", function() { return tableTitleInvisibleTextZIndex; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableTitleUnderlineOverlayZIndex", function() { return tableTitleUnderlineOverlayZIndex; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableTitleInputElementZIndex", function() { return tableTitleInputElementZIndex; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "suggestionsMenuZIndex", function() { return suggestionsMenuZIndex; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "personaCardZIndex", function() { return personaCardZIndex; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "personaBubbleZIndex", function() { return personaBubbleZIndex; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "dragPreviewZIndex", function() { return dragPreviewZIndex; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tableGrabberMenuZIndex", function() { return tableGrabberMenuZIndex; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "selectedCellZIndex", function() { return selectedCellZIndex; });
// ZIndex of the addRowColumnButton so it shows up above other table content.
const addRowColumnButtonZIndex = 10;
// ZIndex of the column resizeHandle, must be on top of the table border to get the right events.
const columnResizeHandleZIndex = 10;
// ZIndex of the invisible text used to provide the right width of the table title underline.
const tableTitleInvisibleTextZIndex = 0;
// ZIndex of the table title underline so that it shows up on top of the default underline.
const tableTitleUnderlineOverlayZIndex = 10;
// ZIndex of the table title input element so that it shows up over the invisible measurement text.
const tableTitleInputElementZIndex = 10;
// ZIndex of the suggestions menu that shows up as the user is typing in the input box.
const suggestionsMenuZIndex = 20;
// ZIndex of the card with persona information when tablero personas are hovered.
const personaCardZIndex = 30;
// ZIndex of the bubble that appears with the coauthor's initials around a cell.
const personaBubbleZIndex = 10;
// ZIndex of preview generated when grabber is dragged.
const dragPreviewZIndex = 20;
// ZIndex of table grabber menu
const tableGrabberMenuZIndex = 100;
// ZIndex of the table cell when selected so that it's shadow is displayed over neighboring cells
const selectedCellZIndex = 1;


/***/ })

}]);
//# sourceMappingURL=fluidTable.js.map