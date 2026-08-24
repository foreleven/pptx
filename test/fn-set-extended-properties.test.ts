// setExtendedProperties — partial setter for /docProps/app.xml.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  getExtendedProperties,
  loadPresentation,
  savePresentation,
  setExtendedProperties,
} from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));
const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

describe('fn API: setExtendedProperties', () => {
  it('updates only the requested fields', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const before = getExtendedProperties(pres)!;
    setExtendedProperties(pres, { company: 'Acme', manager: 'Dana' });

    const after = getExtendedProperties(pres)!;
    expect(after.company).toBe('Acme');
    expect(after.manager).toBe('Dana');
    // Untouched fields are preserved verbatim.
    expect(after.application).toBe(before.application);
    expect(after.appVersion).toBe(before.appVersion);
    expect(after.presentationFormat).toBe(before.presentationFormat);
  });

  it('persists through save → reload', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    setExtendedProperties(pres, {
      company: 'Acme',
      hyperlinkBase: 'https://acme.example/',
    });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const props = getExtendedProperties(reloaded)!;
    expect(props.company).toBe('Acme');
    expect(props.hyperlinkBase).toBe('https://acme.example/');
  });

  it('clears a field when passed null', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    setExtendedProperties(pres, { application: null });
    expect(getExtendedProperties(pres)!.application).toBeNull();
  });

  it('writes presentation counts, heading pairs, and titles of parts', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    setExtendedProperties(pres, {
      slides: 2,
      notes: 1,
      hiddenSlides: 1,
      headingPairs: [
        { heading: 'Theme', count: 1 },
        { heading: 'Slide Titles', count: 2 },
      ],
      titlesOfParts: ['Office Theme', 'Overview', 'Appendix'],
    });

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getExtendedProperties(reloaded)).toMatchObject({
      slides: 2,
      notes: 1,
      hiddenSlides: 1,
      headingPairs: [
        { heading: 'Theme', count: 1 },
        { heading: 'Slide Titles', count: 2 },
      ],
      titlesOfParts: ['Office Theme', 'Overview', 'Appendix'],
    });
  });

  skipIfNoXmllint(
    'removes nullable numeric and vector properties without invalid empty elements',
    async () => {
      const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
      setExtendedProperties(pres, {
        slides: null,
        notes: null,
        hiddenSlides: null,
        headingPairs: null,
        titlesOfParts: null,
      });

      const bytes = await savePresentation(pres);
      const reloaded = await loadPresentation(bytes);
      expect(getExtendedProperties(reloaded)).toMatchObject({
        slides: null,
        notes: null,
        hiddenSlides: null,
        headingPairs: null,
        titlesOfParts: null,
      });
      const appXml = unzipSync(bytes)['docProps/app.xml'];
      if (!appXml) throw new Error('saved presentation has no docProps/app.xml');
      expectSchemaValid(new TextDecoder().decode(appXml), 'extendedProperties');
    },
  );
});
