/* Microsoft PowerPoint host adapter (Office.js PowerPointApi), inspired by Claude in PowerPoint. */

import {
  parseSlideSpec,
  matchSlideLayout,
  classifyPlaceholders,
  formatPresentationSnapshot,
  parseSlideRef,
} from '../ppt-utils.js';
import { parseCellGrid } from '../excel-utils.js';

export const PPT_QUICK_ACTIONS = [
  { cmd: '/review-deck', title: 'Review full deck', prompt: 'Review the entire presentation for narrative flow, pyramid-principle clarity, slide density, inconsistencies, and missing evidence. Cite specific slides (e.g. `Slide 2`) before proposing edits.' },
  { cmd: '/tighten-slide', title: 'Tighten current slide', prompt: 'Tighten the text on the selected slide or shape into crisp, scannable executive bullets (3–5 bullets, concise wording) using update_shape.' },
  { cmd: '/rewrite-title', title: 'Action takeaway title', prompt: 'Rewrite the title on the selected slide into a sharp, insight-led action takeaway title using update_shape.' },
  { cmd: '/add-slide', title: 'Add on-brand slide', prompt: 'Draft and insert a new slide that fits naturally into this deck using one of the available Slide Master layouts via add_slide.' },
  { cmd: '/outline-deck', title: 'Outline / expand deck', prompt: 'Propose a structured slide-by-slide outline to strengthen or extend this presentation, specifying the recommended layout and key message for each slide.' },
  { cmd: '/speaker-notes', title: 'Draft speaker notes', prompt: 'Write concise, natural talking points / speaker notes for the selected slide (or deck) — clear narrative transitions, not reading bullets verbatim.' },
  { cmd: '/add-table', title: 'Insert slide table', prompt: 'Organize the key comparison or data points into a clean native PowerPoint table on this slide using add_table.' },
  { cmd: '/polish-copy', title: 'Proofread whole deck', prompt: 'Proofread the entire deck for typos, awkward phrasing, inconsistent capitalization, and non-parallel bullets. Propose targeted fixes with replace_slide_text.' },
  { cmd: '/zh', title: 'Translate slide to Chinese', prompt: 'Translate the selected shape or slide copy into natural, executive Chinese using update_shape.' },
  { cmd: '/en', title: 'Translate slide to English', prompt: 'Translate the selected shape or slide copy into natural, executive English using update_shape.' },
];

const NON_TEXT_SHAPE_TYPES = new Set(['Image', 'Picture', 'Line', 'Chart', 'SmartArt', 'Media', 'Ink', 'Table', 'Unsupported']);

function cleanSlideMarkdown(text) {
  return String(text || '')
    .replace(/\*\*([^*\n]+)\*\*/g, '$1')
    .replace(/__([^_\n]+)__/g, '$1')
    .replace(/(^|[^*])\*([^*\n]+)\*(?!\*)/g, '$1$2');
}

function applySlideFontSafe(shape, { font, color, fontSize, bold } = {}) {
  try {
    const f = shape?.textFrame?.textRange?.font;
    if (!f) return;
    if (font) f.name = font;
    if (color) f.color = color;
    if (fontSize && Number(fontSize) > 0) f.size = Number(fontSize);
    if (bold) f.bold = true;
  } catch {
    // Ignore unsupported font properties on older PowerPoint builds.
  }
}

function canHaveTextFrame(shape) {
  if (!shape) return false;
  if (shape.type && NON_TEXT_SHAPE_TYPES.has(String(shape.type))) return false;
  return Boolean(shape.textFrame);
}

async function readLayoutsSafe(ctx) {
  const layouts = [];
  try {
    if (!ctx.presentation.slideMasters) return layouts;
    const masters = ctx.presentation.slideMasters;
    masters.load('items/id,items/name');
    await ctx.sync();
    for (const m of masters.items || []) {
      m.layouts.load('items/id,items/name');
    }
    await ctx.sync();
    for (const m of masters.items || []) {
      for (const l of m.layouts.items || []) {
        layouts.push({
          slideMasterId: m.id,
          slideMasterName: m.name || '',
          layoutId: l.id,
          layoutName: l.name || '',
        });
      }
    }
  } catch {
    // SlideMasters API unavailable on older PowerPoint builds.
  }
  return layouts;
}

async function readSelectedSlideIdsSafe(ctx, Office) {
  // PowerPointApi 1.5+
  try {
    if (typeof ctx.presentation.getSelectedSlides === 'function') {
      const selSlides = ctx.presentation.getSelectedSlides();
      selSlides.load('items/id');
      await ctx.sync();
      if (selSlides.items && selSlides.items.length > 0) {
        return selSlides.items.map(s => s.id);
      }
    }
  } catch {
    // Fall back to Common API below
  }

  // Fallback: Common API SlideRange
  if (Office?.context?.document?.getSelectedDataAsync && Office?.CoercionType?.SlideRange) {
    return new Promise(resolve => {
      try {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, res => {
          if (res.status === Office.AsyncResultStatus.Succeeded && res.value?.slides) {
            resolve(res.value.slides.map(s => String(s.id)));
          } else {
            resolve([]);
          }
        });
      } catch {
        resolve([]);
      }
    });
  }
  return [];
}

async function readSelectedShapesSafe(ctx) {
  try {
    if (typeof ctx.presentation.getSelectedShapes === 'function') {
      const selShapes = ctx.presentation.getSelectedShapes();
      selShapes.load('items/id,items/name,items/type');
      await ctx.sync();
      const items = selShapes.items || [];
      for (const s of items) {
        if (canHaveTextFrame(s)) {
          try {
            s.textFrame.load('hasText');
            s.textFrame.textRange.load('text');
          } catch {}
        }
      }
      if (items.length) {
        try { await ctx.sync(); } catch {}
      }
      return items.map(s => ({
        id: s.id,
        name: s.name || '',
        type: s.type || '',
        text: s.textFrame?.hasText ? (s.textFrame.textRange?.text || '') : '',
      }));
    }
  } catch {
    // getSelectedShapes requires PowerPointApi 1.5+
  }
  return [];
}

async function loadShapeTexts(ctx, shapes) {
  const candidates = (shapes || []).filter(canHaveTextFrame);
  for (const s of candidates) {
    try {
      s.textFrame.load('hasText');
      s.textFrame.textRange.load('text');
    } catch {}
  }
  if (!candidates.length) return;
  try {
    await ctx.sync();
  } catch {
    // Some slides contain hybrid shapes that reject batched textFrame load; retry individually.
    for (const s of candidates) {
      try {
        s.textFrame.load('hasText');
        s.textFrame.textRange.load('text');
        await ctx.sync();
      } catch {}
    }
  }
}

function resolveTargetSlide(slidesItems, slideAttr, selectedSlideIds) {
  const slideNum = parseSlideRef(slideAttr);
  if (slideNum !== null) {
    const target = slidesItems[slideNum - 1];
    if (!target) throw new Error(`Slide ${slideNum} does not exist (deck has ${slidesItems.length} slides).`);
    return { slide: target, index: slideNum };
  }
  if (selectedSlideIds.length > 0) {
    const idx = slidesItems.findIndex(s => String(s.id) === String(selectedSlideIds[0]));
    if (idx >= 0) return { slide: slidesItems[idx], index: idx + 1 };
  }
  if (slidesItems.length > 0) {
    return { slide: slidesItems[0], index: 1 };
  }
  throw new Error('This presentation has no slides yet. Use Add slide first.');
}

function findShapeOnSlide(shapes, shapeQuery, selectedShapeIds) {
  if (shapeQuery) {
    const q = String(shapeQuery).trim().toLowerCase();
    const byId = shapes.find(s => String(s.id).toLowerCase() === q);
    if (byId) return byId;
    const exactName = shapes.find(s => String(s.name || '').trim().toLowerCase() === q);
    if (exactName) return exactName;
    const subName = shapes.find(s => String(s.name || '').toLowerCase().includes(q));
    if (subName) return subName;
  }
  if (selectedShapeIds.length > 0) {
    const selected = shapes.find(s => selectedShapeIds.includes(s.id));
    if (selected) return selected;
  }
  // Fallback to first text-capable shape if shapeQuery indicates title/body
  if (shapeQuery && /title/i.test(shapeQuery)) {
    const { titleShape } = classifyPlaceholders(shapes);
    if (titleShape) return titleShape;
  }
  return null;
}

export function createPowerPointAdapter(env = globalThis) {
  const getPpt = () => env.PowerPoint;
  const getOffice = () => env.Office;

  return {
    id: 'powerpoint',
    label: 'PowerPoint',
    noun: 'presentation',
    placeholder: 'Ask about your slides, refine copy, or type / for deck actions…',
    quickActions: PPT_QUICK_ACTIONS,
    secondaryToggle: null,
    modeLabels: {
      auto: 'Context: Auto (whole deck + current slide)',
      selection: 'Context: Current slide / shape only',
      document: 'Context: Whole deck',
    },

    formatContextBadge(ctx, contextMode) {
      const meta = ctx.hostContext || {};
      const slideCount = meta.slideCount || 0;
      const curSlide = meta.currentSlideIndex || 0;
      const shapeName = meta.selectedShapeName || '';
      const layoutCount = Array.isArray(meta.layouts) ? meta.layouts.length : 0;

      if (contextMode === 'selection') {
        if (shapeName) {
          return {
            badge: `Slide ${curSlide} · ${shapeName}`,
            preview: (ctx.selectionText || '').replace(/\s+/g, ' ').trim().slice(0, 80) || 'Shape selected',
            hasSelection: true,
          };
        }
        return {
          badge: curSlide ? `Slide ${curSlide} only` : 'Slide only',
          preview: curSlide ? `Focused on Slide ${curSlide} of ${slideCount}` : 'Select a slide in PowerPoint',
          hasSelection: Boolean(curSlide),
        };
      }
      if (contextMode === 'document') {
        return {
          badge: `Deck · ${slideCount} slide${slideCount === 1 ? '' : 's'}`,
          preview: layoutCount ? `Full deck + ${layoutCount} template layouts` : 'Full presentation attached',
          hasSelection: false,
        };
      }
      if (shapeName) {
        return {
          badge: `Slide ${curSlide} · ${shapeName}`,
          preview: (ctx.selectionText || '').replace(/\s+/g, ' ').trim().slice(0, 80) || 'Shape selected',
          hasSelection: true,
        };
      }
      if (curSlide) {
        return {
          badge: `Slide ${curSlide} / ${slideCount}`,
          preview: meta.currentSlideTitle || (layoutCount ? `Deck attached · ${layoutCount} layouts` : 'Deck attached'),
          hasSelection: true,
        };
      }
      return {
        badge: slideCount ? `Deck · ${slideCount} slides` : 'Presentation',
        preview: slideCount ? 'Full deck attached' : 'Add a slide to begin',
        hasSelection: false,
      };
    },

    getStarters(ctx) {
      const meta = ctx.hostContext || {};
      const curSlide = meta.currentSlideIndex || 0;
      const slideCount = meta.slideCount || 0;
      const hasShape = Boolean(meta.selectedShapeName);

      const starters = hasShape || curSlide
        ? [
            { title: `Tighten Slide ${curSlide || 1}`, sub: 'Crisp executive bullets, cut clutter', prompt: `Tighten the copy on Slide ${curSlide || 1} into crisp, scannable executive bullets using update_shape.` },
            { title: 'Write an action takeaway title', sub: 'Turn the headline into the core insight', prompt: `Rewrite the title of Slide ${curSlide || 1} into a sharp, insight-led takeaway title using update_shape.` },
            { title: 'Review full deck narrative', sub: 'Flow, gaps, and consistency across all slides', prompt: 'Review the entire presentation for narrative flow, slide density, and clarity. Cite specific slides before making edits.' },
            { title: 'Add an on-brand slide', sub: 'Matches your Slide Master layouts & fonts', prompt: 'Draft and add a new summary or next-steps slide that uses an existing layout from this deck\'s Slide Master via add_slide.' },
          ]
        : [
            { title: 'Review deck narrative & flow', sub: 'Honest executive critique slide by slide', prompt: 'Review the entire presentation for story arc, clarity, and slide density. Cite specific slides before making edits.' },
            { title: 'Turn titles into takeaways', sub: 'Pyramid-principle headlines across the deck', prompt: 'Review all slide titles in this deck and propose sharper, takeaway-driven titles.' },
            { title: 'Add a new on-brand slide', sub: 'Built directly on your template layouts', prompt: 'Draft and insert a clean new slide matching the deck\'s Slide Master layouts using add_slide.' },
            { title: 'Proofread the whole deck', sub: 'Fix typos and parallel bullet phrasing', prompt: 'Proofread the entire deck for typos, consistency, and parallel phrasing, proposing targeted edits with replace_slide_text.' },
          ];

      const subtitle = hasShape
        ? `Selected ${meta.selectedShapeName} on Slide ${curSlide} · Ask to rewrite, tighten, or add slides.`
        : slideCount
          ? `Reading ${slideCount} slide${slideCount === 1 ? '' : 's'}${curSlide ? ` (on Slide ${curSlide})` : ''} · Ask anything or type / for actions.`
          : 'Chat about your deck, or ask to create on-brand slides directly.';

      return { heading: 'What are we building?', subtitle, starters };
    },

    async readContext() {
      const PowerPoint = getPpt();
      const Office = getOffice();
      if (!PowerPoint || !PowerPoint.run) return null;

      let result = null;
      await PowerPoint.run(async ctx => {
        const slidesCol = ctx.presentation.slides;
        slidesCol.load('items/id');
        await ctx.sync();

        const slideItems = slidesCol.items || [];
        const [layouts, selectedSlideIds, selectedShapes] = await Promise.all([
          readLayoutsSafe(ctx),
          readSelectedSlideIdsSafe(ctx, Office),
          readSelectedShapesSafe(ctx),
        ]);

        // On huge decks (>60 slides), window around the active slide + first/last slides to avoid Office.js payload timeouts
        const MAX_LOADED_SLIDES = 60;
        const selectedZeroIdx = Math.max(0, slideItems.findIndex(s => selectedSlideIds.includes(s.id)));
        let indicesToLoad;
        if (slideItems.length <= MAX_LOADED_SLIDES) {
          indicesToLoad = slideItems.map((_, i) => i);
        } else {
          const picked = new Set([0, slideItems.length - 1]);
          const half = Math.floor((MAX_LOADED_SLIDES - 2) / 2);
          const start = Math.max(0, Math.min(slideItems.length - MAX_LOADED_SLIDES, selectedZeroIdx - half));
          for (let i = start; i < slideItems.length && picked.size < MAX_LOADED_SLIDES; i++) {
            picked.add(i);
          }
          indicesToLoad = [...picked].sort((a, b) => a - b);
        }
        const loadedSlides = indicesToLoad.map(i => ({ slide: slideItems[i], index: i + 1 }));
        const slidesOmitted = slideItems.length - loadedSlides.length;

        // Load shapes & layout name per windowed slide
        for (const { slide } of loadedSlides) {
          slide.shapes.load('items/id,items/name,items/type,items/left,items/top,items/width,items/height');
          try {
            if (slide.layout) slide.layout.load('name');
          } catch {}
        }
        if (loadedSlides.length) await ctx.sync();

        // Load shape text frames across windowed slides
        const allShapes = loadedSlides.flatMap(({ slide }) => slide.shapes.items || []);
        await loadShapeTexts(ctx, allShapes);

        // Extract text from Grouped shapes and native PowerPoint Tables safely when supported
        const extraTextByShapeId = new Map();
        const groupShapes = allShapes.filter(sh => sh.type === 'Group' && sh.group?.shapes);
        if (groupShapes.length) {
          try {
            for (const g of groupShapes) {
              g.group.shapes.load('items/id,items/name,items/type');
            }
            await ctx.sync();
            const children = groupShapes.flatMap(g => g.group.shapes.items || []);
            await loadShapeTexts(ctx, children);
            for (const g of groupShapes) {
              const texts = (g.group.shapes.items || [])
                .map(c => (c.textFrame?.hasText ? (c.textFrame.textRange?.text || '').trim() : ''))
                .filter(Boolean);
              if (texts.length) extraTextByShapeId.set(g.id, texts.join('\n'));
            }
          } catch {}
        }

        const tableShapes = allShapes.filter(sh => sh.type === 'Table' && typeof sh.getTable === 'function');
        if (tableShapes.length) {
          try {
            const tables = tableShapes.map(sh => {
              const t = sh.getTable();
              t.load('values');
              return { id: sh.id, table: t };
            });
            await ctx.sync();
            for (const { id, table } of tables) {
              if (Array.isArray(table.values) && table.values.length) {
                const formatted = table.values
                  .slice(0, 15)
                  .map(row => `| ${row.map(c => String(c ?? '').replace(/\r?\n/g, ' ')).join(' | ')} |`)
                  .join('\n');
                extraTextByShapeId.set(id, formatted);
              }
            }
          } catch {}
        }

        const slidesData = loadedSlides.map(({ slide, index }) => ({
          index,
          id: slide.id,
          layoutName: slide.layout?.name || '',
          shapes: (slide.shapes.items || []).map(sh => ({
            id: sh.id,
            name: sh.name || '',
            type: sh.type || '',
            left: sh.left,
            top: sh.top,
            width: sh.width,
            height: sh.height,
            text: extraTextByShapeId.get(sh.id) || (sh.textFrame?.hasText ? (sh.textFrame.textRange?.text || '') : ''),
          })),
        }));

        const selectedSlideIndices = slidesData
          .filter(s => selectedSlideIds.includes(s.id))
          .map(s => s.index);

        const currentSlide = selectedSlideIndices.length
          ? slidesData.find(s => s.index === selectedSlideIndices[0])
          : slidesData[0] || null;

        const docText = formatPresentationSnapshot({
          slides: slidesData,
          layouts,
          selectedSlideIndices,
        });

        let selectionText = '';
        let selectedShapeName = '';
        let selectedShapeText = '';

        if (selectedShapes.length > 0) {
          selectedShapeName = selectedShapes[0].name || selectedShapes[0].type || 'Shape';
          selectedShapeText = selectedShapes[0].text || '';
          selectionText = selectedShapes
            .map(s => `Slide ${currentSlide?.index || 1} · [${s.name || s.id}]\n${s.text || '(no text)'}`)
            .join('\n\n');
        } else if (currentSlide) {
          selectionText = formatPresentationSnapshot({
            slides: [currentSlide],
            layouts: [],
            selectedSlideIndices: [currentSlide.index],
          });
        }

        const currentTitleShape = currentSlide ? classifyPlaceholders(currentSlide.shapes).titleShape : null;

        result = {
          selectionText,
          documentText: docText,
          comments: [],
          hostContext: {
            slideCount: slideItems.length,
            slidesOmitted,
            currentSlideIndex: currentSlide?.index || 0,
            currentSlideId: currentSlide?.id || '',
            currentSlideTitle: (currentTitleShape?.text || '').replace(/\s+/g, ' ').trim().slice(0, 70),
            selectedShapeName,
            selectedShapeText,
            layouts,
            slides: slidesData.map(s => ({
              index: s.index,
              id: s.id,
              shapes: s.shapes.map(sh => ({ id: sh.id, name: sh.name, text: sh.text })),
            })),
          },
        };
      });
      return result;
    },

    watchSelection(onChange, onFallback) {
      const Office = getOffice();
      try {
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onChange, result => {
          if (result.status !== Office.AsyncResultStatus.Succeeded && onFallback) {
            onFallback('Auto-selection unavailable here. Use the refresh icon after switching slides.');
          }
        });
      } catch {
        if (onFallback) onFallback('Auto-selection unavailable here. Use the refresh icon after switching slides.');
      }
    },

    async navigateToRef(ref) {
      const PowerPoint = getPpt();
      const Office = getOffice();
      const slideNum = parseSlideRef(ref);
      if (slideNum === null) return { ok: false, message: `Invalid slide reference: ${ref}` };
      if (!PowerPoint || !PowerPoint.run) {
        return { ok: false, message: 'Open inside Microsoft PowerPoint to jump to slides.' };
      }

      let slideId = '';
      await PowerPoint.run(async ctx => {
        const slides = ctx.presentation.slides;
        slides.load('items/id');
        await ctx.sync();
        const target = (slides.items || [])[slideNum - 1];
        if (!target) throw new Error(`Slide ${slideNum} does not exist.`);
        slideId = target.id;
        if (typeof ctx.presentation.setSelectedSlides === 'function') {
          ctx.presentation.setSelectedSlides([target.id]);
          await ctx.sync();
        }
      });

      if (slideId && Office?.context?.document?.goToByIdAsync && Office?.GoToType?.Slide) {
        await new Promise(resolve => {
          try {
            Office.context.document.goToByIdAsync(slideId, Office.GoToType.Slide, () => resolve());
          } catch {
            resolve();
          }
        });
      }

      return { ok: true, message: `Jumped to Slide ${slideNum}` };
    },

    async applyEdit(edit, { brandTheme = null } = {}) {
      const PowerPoint = getPpt();
      const Office = getOffice();
      if (!PowerPoint || !PowerPoint.run) {
        return { ok: false, message: 'Open inside Microsoft PowerPoint to apply edits directly.' };
      }

      const action = edit.action;
      const attrs = edit.attrs || {};
      const content = (edit.content || '').trim();
      let summaryMsg = 'Applied to PowerPoint · Cmd+Z to undo';

      await PowerPoint.run(async ctx => {
        const slidesCol = ctx.presentation.slides;
        slidesCol.load('items/id');
        await ctx.sync();
        const slideItems = slidesCol.items || [];
        const selectedSlideIds = await readSelectedSlideIdsSafe(ctx, Office);

        if (action === 'add_slide') {
          const spec = parseSlideSpec(content);
          const layouts = await readLayoutsSafe(ctx);
          const matchedLayout = matchSlideLayout(layouts, attrs.layout || '');

          const addOpts = {};
          if (matchedLayout?.slideMasterId) addOpts.slideMasterId = matchedLayout.slideMasterId;
          if (matchedLayout?.layoutId) addOpts.layoutId = matchedLayout.layoutId;

          const afterRaw = String(attrs.after_slide ?? '').trim();
          const afterNum = afterRaw === '0' ? 0 : parseSlideRef(afterRaw);
          if (afterNum !== null && afterNum >= 0 && afterNum <= slideItems.length) {
            addOpts.index = afterNum;
          }

          if (Object.keys(addOpts).length > 0) {
            try {
              slidesCol.add(addOpts);
            } catch {
              // Fallback without index if PowerPoint build does not yet support AddSlideOptions.index
              const safeOpts = {};
              if (addOpts.slideMasterId) safeOpts.slideMasterId = addOpts.slideMasterId;
              if (addOpts.layoutId) safeOpts.layoutId = addOpts.layoutId;
              slidesCol.add(Object.keys(safeOpts).length ? safeOpts : undefined);
            }
          } else {
            slidesCol.add();
          }
          await ctx.sync();

          // Locate the newly added slide
          slidesCol.load('items/id');
          await ctx.sync();
          const prevIds = new Set(slideItems.map(s => s.id));
          const newSlide = (slidesCol.items || []).find(s => !prevIds.has(s.id)) ||
            slidesCol.items[slidesCol.items.length - 1];

          if (newSlide) {
            newSlide.shapes.load('items/id,items/name,items/type,items/left,items/top,items/width,items/height');
            await ctx.sync();
            const placeholders = classifyPlaceholders(newSlide.shapes.items || []);

            const bt = brandTheme && brandTheme.enabled ? brandTheme : null;
            const slideFont = (attrs.font || '').trim() || (bt?.font || '');
            const slideColor = (attrs.color || '').trim() || (bt?.color || '');
            const titleColor = (attrs.title_color || attrs.titleColor || '').trim() || (bt?.headingColor || '');

            if (spec.title) {
              const cleanTitle = cleanSlideMarkdown(spec.title);
              if (placeholders.titleShape && placeholders.titleShape.textFrame) {
                placeholders.titleShape.textFrame.textRange.text = cleanTitle;
                applySlideFontSafe(placeholders.titleShape, { font: slideFont, color: titleColor || slideColor });
              } else if (typeof newSlide.shapes.addTextBox === 'function') {
                const tb = newSlide.shapes.addTextBox(cleanTitle, { left: 48, top: 36, width: 624, height: 64 });
                applySlideFontSafe(tb, { font: slideFont, color: titleColor || slideColor, fontSize: 28, bold: true });
              }
            }

            if (spec.subtitle) {
              const cleanSub = cleanSlideMarkdown(spec.subtitle);
              if (placeholders.subtitleShape && placeholders.subtitleShape.textFrame) {
                placeholders.subtitleShape.textFrame.textRange.text = cleanSub;
                applySlideFontSafe(placeholders.subtitleShape, { font: slideFont, color: slideColor });
              } else if (placeholders.bodyShapes[0] && !spec.body && !spec.columns.length) {
                placeholders.bodyShapes[0].textFrame.textRange.text = cleanSub;
                applySlideFontSafe(placeholders.bodyShapes[0], { font: slideFont, color: slideColor });
              }
            }

            if (spec.columns.length > 0) {
              spec.columns.forEach((colText, idx) => {
                const cleanCol = cleanSlideMarkdown(colText);
                const target = placeholders.bodyShapes[idx];
                if (target && target.textFrame) {
                  target.textFrame.textRange.text = cleanCol;
                  applySlideFontSafe(target, { font: slideFont, color: slideColor });
                } else if (typeof newSlide.shapes.addTextBox === 'function') {
                  const colW = Math.floor(600 / spec.columns.length) - 12;
                  const tb = newSlide.shapes.addTextBox(cleanCol, {
                    left: 48 + idx * (colW + 24),
                    top: 120,
                    width: colW,
                    height: 320,
                  });
                  applySlideFontSafe(tb, { font: slideFont, color: slideColor });
                }
              });
            } else if (spec.body) {
              const cleanBody = cleanSlideMarkdown(spec.body);
              const targetBody = placeholders.bodyShapes[0] || (!spec.subtitle ? placeholders.subtitleShape : null);
              if (targetBody && targetBody.textFrame) {
                targetBody.textFrame.textRange.text = cleanBody;
                applySlideFontSafe(targetBody, { font: slideFont, color: slideColor });
              } else if (typeof newSlide.shapes.addTextBox === 'function') {
                const tb = newSlide.shapes.addTextBox(cleanBody, { left: 48, top: 120, width: 624, height: 320 });
                applySlideFontSafe(tb, { font: slideFont, color: slideColor });
              }
            }

            if (typeof ctx.presentation.setSelectedSlides === 'function') {
              ctx.presentation.setSelectedSlides([newSlide.id]);
            }
            await ctx.sync();
          }

          summaryMsg = `Inserted slide${matchedLayout?.layoutName ? ` (${matchedLayout.layoutName})` : ''} · Cmd+Z to undo`;
          return;
        }

        if (action === 'delete_slide') {
          const { slide, index } = resolveTargetSlide(slideItems, attrs.slide, selectedSlideIds);
          slide.delete();
          await ctx.sync();
          summaryMsg = `Deleted Slide ${index} · Cmd+Z to undo`;
          return;
        }

        if (action === 'add_textbox') {
          if (!content) throw new Error('Text box content is empty.');
          const { slide, index } = resolveTargetSlide(slideItems, attrs.slide, selectedSlideIds);
          if (typeof slide.shapes.addTextBox !== 'function') {
            throw new Error('This version of PowerPoint does not support adding text boxes from add-ins.');
          }
          const opts = {
            left: Number(attrs.left) || 60,
            top: Number(attrs.top) || 140,
            width: Number(attrs.width) || 600,
            height: Number(attrs.height) || 120,
          };
          const shape = slide.shapes.addTextBox(cleanSlideMarkdown(content), opts);
          applySlideFontSafe(shape, {
            font: (attrs.font || '').trim(),
            color: (attrs.color || '').trim(),
            fontSize: attrs.font_size,
            bold: String(attrs.bold || '').toLowerCase() === 'true',
          });
          await ctx.sync();
          summaryMsg = `Added text box to Slide ${index} · Cmd+Z to undo`;
          return;
        }

        if (action === 'add_table') {
          const grid = parseCellGrid(content);
          if (!grid.length || !grid[0].length) {
            throw new Error('Could not parse table rows.');
          }
          const { slide, index } = resolveTargetSlide(slideItems, attrs.slide, selectedSlideIds);
          if (typeof slide.shapes.addTable !== 'function') {
            throw new Error('This version of PowerPoint does not support adding native tables from add-ins (requires PowerPointApi 1.8+).');
          }
          const strGrid = grid.map(row => row.map(cell => cleanSlideMarkdown(cell ?? '')));
          slide.shapes.addTable(strGrid.length, strGrid[0].length, {
            values: strGrid,
            left: Number(attrs.left) || 48,
            top: Number(attrs.top) || 130,
            width: Number(attrs.width) || 624,
            height: Number(attrs.height) || Math.min(340, strGrid.length * 42 + 24),
          });
          await ctx.sync();
          summaryMsg = `Inserted ${strGrid.length}×${strGrid[0].length} table on Slide ${index} · Cmd+Z to undo`;
          return;
        }

        if (action === 'update_shape') {
          if (!content) throw new Error('New shape text is empty.');
          const { slide, index } = resolveTargetSlide(slideItems, attrs.slide, selectedSlideIds);
          const selectedShapes = await readSelectedShapesSafe(ctx);
          const selectedShapeIds = selectedShapes.map(s => s.id);

          slide.shapes.load('items/id,items/name,items/type,items/left,items/top');
          await ctx.sync();

          const target = findShapeOnSlide(slide.shapes.items || [], attrs.shape, selectedShapeIds);
          if (!target) {
            throw new Error(`Could not find shape${attrs.shape ? ` "${attrs.shape}"` : ''} on Slide ${index}. Select the text box in PowerPoint and retry.`);
          }
          if (!target.textFrame || !target.textFrame.textRange) {
            throw new Error(`Shape "${target.name || target.id}" does not support text editing.`);
          }
          target.textFrame.textRange.text = cleanSlideMarkdown(content);
          await ctx.sync();
          summaryMsg = `Updated ${target.name || 'shape'} on Slide ${index} · Cmd+Z to undo`;
          return;
        }

        if (action === 'replace_slide_text') {
          const findStr = String(edit.find || '').trim();
          if (!findStr) throw new Error('Missing <find> text to replace.');
          const scopeSlides = attrs.slide
            ? [resolveTargetSlide(slideItems, attrs.slide, selectedSlideIds).slide]
            : slideItems;

          for (const sl of scopeSlides) {
            sl.shapes.load('items/id,items/name,items/type');
          }
          await ctx.sync();

          const allShapes = scopeSlides.flatMap(sl => sl.shapes.items || []);
          await loadShapeTexts(ctx, allShapes);

          let replacedCount = 0;
          for (const sh of allShapes) {
            if (!canHaveTextFrame(sh) || !sh.textFrame?.hasText) continue;
            const current = sh.textFrame.textRange?.text || '';
            if (current.includes(findStr)) {
              sh.textFrame.textRange.text = current.split(findStr).join(content);
              replacedCount++;
            }
          }
          if (replacedCount === 0) {
            throw new Error(`Could not find "${findStr.slice(0, 60)}" on the target slide(s).`);
          }
          await ctx.sync();
          summaryMsg = `Updated text in ${replacedCount} shape${replacedCount === 1 ? '' : 's'} · Cmd+Z to undo`;
          return;
        }

        throw new Error(`Unsupported PowerPoint action: ${action}`);
      });

      return { ok: true, message: summaryMsg };
    },
  };
}
