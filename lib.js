'use strict';

Object.defineProperty(exports, '__esModule', {
  value: true
});
exports.color2argb = color2argb;
exports.size2pt = size2pt;
exports.size2px = size2px;
exports.css2style = css2style;
exports.getBorder = getBorder;

const _tinycolor = require('tinycolor2');

const _tinycolor2 = _interopRequireDefault(_tinycolor);

const _cssstyle = require('cssstyle');

function _interopRequireDefault (obj) { return obj && obj.__esModule ? obj : { default: obj }; }

function color2argb (c) {
  const rgba = (0, _tinycolor2.default)(c).toHex8();
  return rgba.substr(6) + rgba.substr(0, 6);
}

function size2pt (s) {
  const num = size2px(s);
  if (num > 0) {
    return num * 72 / 96;
  }
  return 12;
}

function size2px (s) {
  if (!s) return 0;

  const pt = s.match(/([.\d]+)pt/i);
  if (pt && pt.length === 2) {
    return parseFloat(pt[1], 10) * 96 / 72;
  }
  const em = s.match(/([.\d]+)em/i);
  if (em && em.length === 2) {
    return parseFloat(em[1], 10) * 16;
  }
  const px = s.match(/([.\d]+)px/i);
  if (px && px.length === 2) {
    return parseFloat(px[1], 10);
  }
  const pe = s.match(/([.\d]+)%/i);
  if (pe && pe.length === 2) {
    return parseFloat(pe[1], 10) / 100 * 16;
  }
  return 0;
}

function css2style () {
  const css = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {};

  const style = new _cssstyle.CSSStyleDeclaration();
  let _iteratorNormalCompletion = true;
  let _didIteratorError = false;
  let _iteratorError;

  try {
    for (let _iterator = Object.keys(css)[Symbol.iterator](), _step; !(_iteratorNormalCompletion = (_step = _iterator.next()).done); _iteratorNormalCompletion = true) {
      const k = _step.value;

      style[k] = css[k];
    }
  } catch (err) {
    _didIteratorError = true;
    _iteratorError = err;
  } finally {
    try {
      if (!_iteratorNormalCompletion && _iterator.return) {
        _iterator.return();
      }
    } finally {
      if (_didIteratorError) {
        throw _iteratorError;
      }
    }
  }

  return style;
}

function getBorder (css, type) {
  let color = css[`border-${type}-color`];
  let style = css[`border-${type}-style`];
  let width = css[`border-${type}-width`];

  if (!color) return null;

  width = size2px(width);
  if (width <= 0) return null;

  color = color2argb(color);

  if (style === 'dashed' || style === 'dotted' || style === 'double') {
    return { style, color };
  }
  style = 'thin';
  if (width >= 3 && width < 5) {
    style = 'medium';
  }
  if (width >= 5) {
    style = 'thick';
  }
  return { style, color };
}
