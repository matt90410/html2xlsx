'use strict';

const _betterXlsx = require('better-xlsx');

const _betterXlsx2 = _interopRequireDefault(_betterXlsx);

const _juice = require('juice');

const _juice2 = _interopRequireDefault(_juice);

const _cheerio = require('cheerio');

const _cheerio2 = _interopRequireDefault(_cheerio);

const _moment = require('moment');

const _moment2 = _interopRequireDefault(_moment);

const _lib = require('./lib');

function _interopRequireDefault (obj) { return obj && obj.__esModule ? obj : { default: obj }; }

module.exports = function (html, callback) {
  const options = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : {};

  _juice2.default.juiceResources(html, options.juice || {}, function (err, text) {
    if (err) return callback(err);

    const file = new _betterXlsx2.default.File();
    const $ = _cheerio2.default.load(text);

    $('table').each(function (ti, table) {
      //  MSB - added ability to name the tabs
      const sheetName = $(table).attr('title') ? $(table).attr('title') : `Sheet${ti + 1}`;
      let sheet = file.addSheet(sheetName);

      let maxW = [];
      let offsets = [];
      $('tr', table).each(function (hi, th) {
        if (offsets[hi] === undefined) {
          offsets[hi] = 0;
        }
        let maxH = 20; // pt
        $('th, td', th).each(function (di, td) {
          let $td = $(td);
          const rs = parseInt($td.attr('rowspan'), 10) || 1;
          const cs = parseInt($td.attr('colspan'), 10) || 1;

          for (let r = 0; r < rs; r++) {
            for (let c = 0; c < cs; c++) {
              sheet.cell(hi + r, offsets[hi] + c);
            }
          }

          const css = (0, _lib.css2style)($td.css());
          const fsize = (0, _lib.size2pt)(css.fontSize);
          // Row Height & Col Width
          if (css.height) {
            const pt = (0, _lib.size2pt)(css.height);
            if (pt > maxH) {
              maxH = pt / rs;
            }
          }
          if (css.width) {
            if (!maxW[di]) {
              maxW[di] = 10;
            }
            const tmp = (0, _lib.size2pt)(css.width) / fsize;
            if (maxW[di] < tmp) {
              maxW[di] = tmp / cs;
            }
          }
          let style = new _betterXlsx2.default.Style();
          // Font
          style.font.color = (0, _lib.color2argb)(css.color || '#000');
          style.font.size = fsize;
          style.font.name = css.fontFamily || 'Verdana';
          style.font.bold = css.fontWeight === 'bold';
          style.font.italic = css.fontStyle === 'italic';
          style.font.underline = css.textDecoration === 'underline';
          // Fill
          const bgColor = css.backgroundColor;
          if (bgColor) {
            style.fill.patternType = 'solid';
            style.fill.fgColor = (0, _lib.color2argb)(bgColor);
          }
          // Border
          const left = (0, _lib.getBorder)(css, 'left');
          if (left) {
            style.border.left = left.style;
            style.border.leftColor = left.color;
          }
          const right = (0, _lib.getBorder)(css, 'right');
          if (right) {
            style.border.right = right.style;
            style.border.rightColor = right.color;
          }
          const top = (0, _lib.getBorder)(css, 'top');
          if (top) {
            style.border.top = top.style;
            style.border.topColor = top.color;
          }
          const bottom = (0, _lib.getBorder)(css, 'bottom');
          if (bottom) {
            style.border.bottom = bottom.style;
            style.border.bottomColor = bottom.color;
          }
          // Align
          const hMap = {
            left: 'left',
            right: 'right',
            center: 'center',
            justify: 'justify'
          };
          if (css.textAlign && hMap[css.textAlign]) {
            style.align.h = hMap[css.textAlign];
          }
          const vMap = {
            top: 'top',
            bottom: 'bottom',
            middle: 'center'
          };
          if (css.verticalAlign && vMap[css.verticalAlign]) {
            style.align.v = vMap[css.verticalAlign];
          }

          //  MSB - added so that text is wrapped inside cells
          style.align.wrapText = 1;

          // Cell
          const cell = sheet.cell(hi, offsets[hi]);
          // Set value type
          const text = $td.text().trim();
          const type = $td.attr('type') || $td.attr('data-type') || '';
          const numFormat = $td.attr('data-num-format') || '';

          switch (type.toLowerCase()) {
            case 'number':
              cell.setNumber(text);
              break;
            case 'money':
              //  MSB added
              cell.setNumber(text);
              cell.numFmt = '£#,##0.00';
              break;
            case 'bool':
              cell.setBool(text === 'true' || text === '1');
              break;
            case 'formula':
              cell.setFormula(text);
              break;
            case 'date':
              cell.setDate((0, _moment2.default)(text).toDate());
              break;
            case 'datetime':
              cell.setDateTime((0, _moment2.default)(text).toDate());
              break;
            default:
              cell.value = text;
          }
          cell.style = style;
          if (numFormat) {
            cell.numFmt = numFormat;
          }

          if (rs > 1) {
            cell.vMerge = rs - 1;
          }
          if (cs > 1) {
            cell.hMerge = cs - 1;
          }

          for (let _r = 0; _r < rs; _r++) {
            if (offsets[hi + _r] === undefined) {
              offsets[hi + _r] = 0;
            }
            offsets[hi + _r] += cs;
          }
        });
        //  MSB removed so that calls height is automatic
        //     sheet.rows[hi].setHeightCM(maxH * 0.03528);
      });
      // Set col width
      for (let i = 0; i < maxW.length; i++) {
        const w = maxW[i];
        if (w) {
          sheet.col(i).width = w;
        }
      }
    });

    callback(null, file);
  });
};
