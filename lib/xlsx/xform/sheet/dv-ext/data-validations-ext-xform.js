const _ = require('../../../../utils/under-dash');
const CompositeXform = require('../../composite-xform');
const Range = require('../../../../doc/range');
const colCache = require('../../../../utils/col-cache');

const DataValidationExtForm = require('./data-validation-ext-xform');
const DataValidationExtXform = require('./data-validation-ext-xform');

function optimiseDataValidations(model) {
  const Worksheet = require('../../../../doc/worksheet');
  const {maxRows} = Worksheet;
  const {maxCols} = Worksheet;
  const {bailOnMaxLimit} = Worksheet;
  // Squeeze alike data validations together into rectangular ranges
  // to reduce file size and speed up Excel load time
  // We do it only for external DVs
  const dvList = _.map(model, (dataValidation, address) => ({
    address,
    dataValidation,
    marked: false,
  }))
    .filter(dv => DataValidationExtXform.isExt(dv.dataValidation))
    .sort((a, b) => _.strcmp(a.address, b.address));
  const dvMap = _.keyBy(dvList, 'address');
  const matchCol = (addr, height, col) => {
    for (let i = 0; i < height; i++) {
      const otherAddress = colCache.encodeAddress(addr.row + i, col);
      if (!model[otherAddress] || !_.isEqual(model[addr.address], model[otherAddress])) {
        return false;
      }
    }
    return true;
  };
  return dvList
    .map(function(dv) {
      if (!dv.marked) {
        const addr = colCache.decodeEx(dv.address);
        if (addr.dimensions) {
          dvMap[addr.dimensions].marked = true;
          return {
            ...dv.dataValidation,
            sqref: dv.address,
          };
        }
        // iterate downwards - finding matching cells
        let height = 1;
        let otherAddress = colCache.encodeAddress(addr.row + height, addr.col);
        while (model[otherAddress] && _.isEqual(dv.dataValidation, model[otherAddress])) {
          height++;
          if (height > maxRows) {
            if (bailOnMaxLimit) {
              throw new Error(`Max row count (${maxRows}) exceeded in data validation optimization`);
            } else {
              break;
            }
          }
          otherAddress = colCache.encodeAddress(addr.row + height, addr.col);
        }
        // iterate rightwards...
        let width = 1;
        while (matchCol(addr, height, addr.col + width)) {
          width++;
          if (width > maxCols) {
            if (bailOnMaxLimit) {
              throw new Error(`Max column count (${maxCols}) exceeded in data validation optimization`);
            } else {
              break;
            }
          }
        }
        // mark all included addresses
        for (let i = 0; i < height; i++) {
          for (let j = 0; j < width; j++) {
            if (i >= maxRows || j >= maxCols) continue;
            otherAddress = colCache.encodeAddress(addr.row + i, addr.col + j);
            dvMap[otherAddress].marked = true;
          }
        }
        if (height > 1 || width > 1) {
          const bottom = addr.row + (height - 1);
          const right = addr.col + (width - 1);
          return {
            ...dv.dataValidation,
            sqref: `${dv.address}:${colCache.encodeAddress(bottom, right)}`,
          };
        }
        return {
          ...dv.dataValidation,
          sqref: dv.address,
        };
      }
      return null;
    })
    .filter(Boolean);
}

class DataValidationsExtXform extends CompositeXform {
  constructor() {
    super();

    this.map = {
      'x14:dataValidation': (this.dataValidation = new DataValidationExtForm()),
    };
  }

  get tag() {
    return 'x14:dataValidations';
  }

  render(xmlStream, model) {
    const optimizedModel = optimiseDataValidations(model);
    xmlStream.openNode(this.tag, {
      count: optimizedModel.length,
      'xmlns:xm': 'http://schemas.microsoft.com/office/excel/2006/main',
    });
    optimizedModel.forEach(value => {
      this.dataValidation.render(xmlStream, value);
    });
    xmlStream.closeNode();
  }

  hasContent(model) {
    if (!model) {
      return false;
    }
    if (model.hasExtContent === undefined) {
      model.hasExtContent = _.some(model, dv => DataValidationExtXform.isExt(dv));
    }
    return model.hasExtContent;
  }

  createNewModel() {
    return {};
  }

  onParserClose(name, parser) {
    // Assign the data validation per cell, like regular
    // data validations
    const list = parser.model._address.split(/\s+/g) || [];
    const Worksheet = require('../../../../doc/worksheet');
    const {maxRows, maxCols, bailOnMaxLimit} = Worksheet;
    delete parser.model._address;
    list.forEach(addr => {
      if (addr.includes(':')) {
        const range = new Range(addr);
        // Truncate range to limits
        const limitedBottom = Math.min(range.bottom, range.top + maxRows - 1);
        const limitedRight = Math.min(range.right, range.left + maxCols - 1);
        if (range.bottom - range.top + 1 > maxRows || range.right - range.left + 1 > maxCols) {
          if (bailOnMaxLimit) {
            throw new Error(`Data validation range ${addr} exceeds maxRows (${maxRows}) or maxCols (${maxCols})`);
          }
          // Truncate
          range.bottom = limitedBottom;
          range.right = limitedRight;
        }
        range.forEachAddress(address => {
          this.model[address] = parser.model;
        });
      } else {
        this.model[addr] = parser.model;
      }
    });
  }
}

module.exports = DataValidationsExtXform;
