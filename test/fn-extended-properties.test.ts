// getExtendedProperties — read /docProps/app.xml.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { getExtendedProperties, loadPresentation } from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const rewriteAppXml = (bytes: Uint8Array, rewrite: (xml: string) => string): Uint8Array => {
  const entries = unzipSync(bytes);
  const app = entries['docProps/app.xml'];
  if (!app) throw new Error('fixture has no docProps/app.xml');
  entries['docProps/app.xml'] = new TextEncoder().encode(rewrite(new TextDecoder().decode(app)));
  return zipSync(entries);
};

describe('fn API: getExtendedProperties', () => {
  it('reads Application + AppVersion from the fixture', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const props = getExtendedProperties(pres);
    expect(props).not.toBeNull();
    // python-pptx-generated fixture ships these values.
    expect(props!.application).toBe('Microsoft Macintosh PowerPoint');
    expect(props!.appVersion).toBe('14.0000');
    expect(props!.presentationFormat).toBe('On-screen Show (4:3)');
  });

  it('reads presentation counts, heading pairs, and titles of parts', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const props = getExtendedProperties(pres);
    expect(props).not.toBeNull();
    expect(props).toMatchObject({
      slides: 0,
      notes: 0,
      hiddenSlides: 0,
      headingPairs: [
        { heading: 'Theme', count: 1 },
        { heading: 'Slide Titles', count: 0 },
      ],
      titlesOfParts: ['Office Theme'],
    });
  });

  it('exposes empty string fields as null', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const props = getExtendedProperties(pres);
    expect(props).not.toBeNull();
    // Fixture ships empty <Manager> / <Company> / <HyperlinkBase>.
    expect(props!.manager).toBeNull();
    expect(props!.company).toBeNull();
    expect(props!.hyperlinkBase).toBeNull();
  });

  it('accepts alternate XML Schema integer lexical forms', async () => {
    const source = await readFile(fixture('two-slides.pptx'));
    const pres = await loadPresentation(
      rewriteAppXml(source, (xml) =>
        xml
          .replace('<Slides>0</Slides>', '<Slides> +01 </Slides>')
          .replace('size="4" baseType="variant"', 'size="+04" baseType="variant"')
          .replace('<vt:i4>1</vt:i4>', '<vt:i4>+01</vt:i4>'),
      ),
    );

    expect(getExtendedProperties(pres)).toMatchObject({
      slides: 1,
      headingPairs: [
        { heading: 'Theme', count: 1 },
        { heading: 'Slide Titles', count: 0 },
      ],
    });
  });

  it.each([
    ['<Slides>0</Slides>', '<Slides>2e0</Slides>', /Slides is not an XML Schema integer/u],
    [
      '<Slides>0</Slides>',
      '<Slides>\u00A0+1\u00A0</Slides>',
      /Slides is not an XML Schema integer/u,
    ],
    ['<Slides>0</Slides>', '<Slides>-1</Slides>', /Slides must be non-negative/u],
    ['size="4" baseType="variant"', 'size="4.0" baseType="variant"', /vector size is not/u],
    ['<vt:i4>1</vt:i4>', '<vt:i4>2147483648</vt:i4>', /HeadingPairs count is outside/u],
  ])('rejects invalid numeric XML: %s', async (from, to, message) => {
    const source = await readFile(fixture('two-slides.pptx'));
    const pres = await loadPresentation(rewriteAppXml(source, (xml) => xml.replace(from, to)));
    expect(() => getExtendedProperties(pres)).toThrow(message);
  });
});
