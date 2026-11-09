import { describe, expect, it } from 'vitest';
import { INDIAN_AXIS_FORMAT } from './chart.handler';

describe('P3.20 chart value axis in lakh / crore', () => {
  it('uses escaped literal commas (Indian grouping), not thousands separators', () => {
    expect(INDIAN_AXIS_FORMAT).toBe('[>=10000000]##\\,##\\,##\\,##0;[>=100000]##\\,##\\,##0;##,##0');
    expect(INDIAN_AXIS_FORMAT.includes('\\,')).toBe(true);
    expect(INDIAN_AXIS_FORMAT.split(';')).toHaveLength(3);
  });
});
