import TwoCellAnchorXform from '../../../../../src/xlsx/xform/drawing/two-cell-anchor-xform.ts';

describe('TwoCellAnchorXform', () => {
  describe('reconcile', () => {
    it('should not throw on null picture', () => {
      const twoCell = new TwoCellAnchorXform();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      expect(() => twoCell.reconcile({picture: null} as any, {})).not.toThrow();
    });
    it('should not throw on null tl', () => {
      const twoCell = new TwoCellAnchorXform();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      expect(() =>
        twoCell.reconcile({br: {col: 1, row: 1}} as any, {})
      ).not.toThrow();
    });
    it('should not throw on null br', () => {
      const twoCell = new TwoCellAnchorXform();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      expect(() =>
        twoCell.reconcile({tl: {col: 1, row: 1}} as any, {})
      ).not.toThrow();
    });
  });
});
