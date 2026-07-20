import { describe, expect, it } from 'vitest';
import { shortenActionPreviewCopy } from '@/utils/actionPreviewCopy';

describe('shortenActionPreviewCopy', () => {
  it('collapses verbose append previews to a short summary', () => {
    const long = `Preview (new rows to be appended under headers Job Title | Company/Shop | Student Name | Student Email | Student Phone | Location | Resume Link | Status | Applied Date): Row 4 -> Sales Associate | ABC Retail | Anita George | anita.george@example.com | "9638527410" | Kochi | https://resume.link/anita | Applied | "15/7/2026, 10:30:00 am" Row 5 -> Field Technician | XYZ Services | Rohit Kumar | rohit.kumar@example.com | "8123456789" | Thiruvananthapuram | https://resume.link/rohit | Interview Scheduled | "16/7/2026, 2:15:00 pm" Exactly what will change if you approve: • Cells affected: Applications!A4:I4 and Applications!A5:I5 • Values to be written (no formulas): - A4: "Sales Associate"`;

    expect(shortenActionPreviewCopy(long)).toBe(
      'new rows to be appended under the existing headers.',
    );
  });

  it('leaves short copy mostly intact', () => {
    expect(shortenActionPreviewCopy('Append 2 rows to Applications.')).toBe(
      'Append 2 rows to Applications.',
    );
  });
});
