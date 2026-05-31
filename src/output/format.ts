/**
 * Text rendering of extraction results (compatibility surface).
 *
 * The actual rendering lives in `output/render.ts`, which also handles csv,
 * table, and markdown. These helpers preserve the original text-only API used
 * by the API entry points.
 */
import { ExtractionResult } from '../engine/extract';
import { ExtractConfig } from '../config';
import { renderText } from './render';

/** Render a single extraction result to the documented text layout. */
export function formatResult(result: ExtractionResult, config: ExtractConfig): string {
  return renderText([result], config);
}

/** Render an ordered list of results, blank-line separated. */
export function formatResults(
  results: ExtractionResult[],
  config: ExtractConfig,
): string {
  return renderText(results, config);
}
