
import testUtils from '../../utils/index.ts';

import copyStyle from '../../../src/utils/copy-style.ts';
const style1 = {
  numFmt: testUtils.styles.numFmts.numFmt1,
  font: testUtils.styles.fonts.broadwayRedOutline20,
  alignment: testUtils.styles.namedAlignments.topLeft,
  border: testUtils.styles.borders.thickRainbow,
  fill: testUtils.styles.fills.redGreenDarkTrellis,
};
const style2 = {
  fill: testUtils.styles.fills.rgbPathGrad,
};

describe('copyStyle', () => {
  it('should copy a style deeply', () => {
    const copied = copyStyle(style1);
    expect(copied).toEqual(style1);
    expect(copied.font).not.toBe(style1.font);
    expect(copied.alignment).not.toBe(style1.alignment);
    expect(copied.border).not.toBe(style1.border);
    expect(copied.fill).not.toBe(style1.fill);

    expect(copyStyle({})).toEqual({});
  });

  it('should copy fill.stops deeply', () => {
    const copied = copyStyle(style2);
    expect(copied.fill.stops).toEqual(style2.fill.stops);
    expect(copied.fill.stops).not.toBe(style2.fill.stops);
    expect(copied.fill.stops[0]).not.toBe(style2.fill.stops[0]);
  });

  it('should return the argument if a falsy value passed', () => {
    expect(copyStyle(null)).toBe(null);
    expect(copyStyle(undefined)).toBe(undefined);
  });
});
