import { describe, expect, it } from 'vitest';
import { partName } from '../src/internal/opc/index.ts';
import {
  _internalPackageOf,
  addBlankSlide,
  addSlideMedia,
  createPresentation,
  getGroupChildren,
  getShapeMedia,
  getShapeMediaKind,
  getSlidePartName,
  getSlideTopLevelShapes,
  getSlides,
  groupShapes,
  inches,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';

const POSTER = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
    'base64',
  ),
);

describe('fn API: embedded audio and video', () => {
  it('authors click-to-play media frames and preserves their bytes through save and load', async () => {
    const presentation = createPresentation();
    const slide = addBlankSlide(presentation);
    const videoBytes = Uint8Array.from([0, 0, 0, 24, 102, 116, 121, 112, 109, 112, 52, 50]);
    const audioBytes = Uint8Array.from([73, 68, 51, 4, 0, 0, 0, 0]);
    const video = addSlideMedia(slide, videoBytes, {
      kind: 'video',
      contentType: 'video/mp4',
      posterBytes: POSTER,
      posterFormat: 'png',
      name: 'Product demo',
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2.25),
    });
    const audio = addSlideMedia(slide, audioBytes, {
      kind: 'audio',
      contentType: 'audio/mpeg',
      posterBytes: POSTER,
      posterFormat: 'png',
      name: 'Narration',
      x: inches(1),
      y: inches(4),
      w: inches(2),
      h: inches(1),
    });
    groupShapes([video, audio], { name: 'Media group' });

    const loaded = await loadPresentation(await savePresentation(presentation));
    const group = getSlideTopLevelShapes(getSlides(loaded)[0]!)[0]!;
    const [loadedVideo, loadedAudio] = getGroupChildren(group);

    expect(getShapeMedia(loadedVideo!)).toMatchObject({
      kind: 'video',
      contentType: 'video/mp4',
      bytes: videoBytes,
      posterContentType: 'image/png',
      posterBytes: POSTER,
      clickToPlay: true,
    });
    expect(getShapeMedia(loadedAudio!)).toMatchObject({
      kind: 'audio',
      contentType: 'audio/mpeg',
      bytes: audioBytes,
      posterContentType: 'image/png',
      posterBytes: POSTER,
      clickToPlay: true,
    });
  });

  it('detects media markup independently from external relationship resolution', () => {
    const presentation = createPresentation();
    const slide = addBlankSlide(presentation);
    const video = addSlideMedia(slide, Uint8Array.from([0, 0, 0, 24]), {
      kind: 'video',
      contentType: 'video/mp4',
      posterBytes: POSTER,
      posterFormat: 'png',
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2.25),
    });
    const pkg = _internalPackageOf(presentation);
    const slidePartName = partName(getSlidePartName(slide));
    const relationships = pkg.getRels(slidePartName)!;
    for (const relationship of relationships.items) {
      if (!/\/(?:video|media)$/u.test(relationship.type)) continue;
      relationship.target = 'https://example.com/video.mp4';
      relationship.targetMode = 'External';
    }
    pkg.setRels(slidePartName, relationships);

    expect(getShapeMediaKind(video)).toBe('video');
    expect(getShapeMedia(video)).toBeNull();
  });
});
