#!/usr/bin/env node
/**
 * CLI entry point for `extract-excel`.
 *
 * Thin wrapper over the API: parse argv, dispatch by command kind, render
 * results, and translate {@link ExtractError}s into clean terminal messages
 * with appropriate exit codes.
 */
import { parse } from './parser/tokenizer';
import { runExtract } from './run';
import { runTests } from './test/runner';
import { renderHelp } from './help/docs';
import { formatTestSummary } from './index';
import { ExtractError } from './errors';

async function main(argv: string[]): Promise<number> {
  // With no arguments, show global help rather than erroring out.
  if (argv.length === 0) {
    process.stdout.write(renderHelp({ kind: 'help', topic: 'global' }) + '\n');
    return 0;
  }

  const command = parse(argv);

  switch (command.kind) {
    case 'help':
      process.stdout.write(renderHelp(command) + '\n');
      return 0;

    case 'test': {
      const report = await runTests(command);
      if (command.mode === 'custom') {
        // Custom tests emit the replayed data; print it, then a status line.
        if (report.output) process.stdout.write(report.output);
        if (report.failed > 0) {
          for (const r of report.results) {
            if (!r.ok) process.stderr.write(`error: ${r.error}\n`);
          }
          return 1;
        }
        return 0;
      }
      process.stdout.write(formatTestSummary(report) + '\n');
      return report.failed > 0 ? 1 : 0;
    }

    case 'extract':
      await runExtract(command);
      return 0;
  }
}

main(process.argv.slice(2))
  .then((code) => {
    process.exitCode = code;
  })
  .catch((err: unknown) => {
    if (err instanceof ExtractError) {
      process.stderr.write(`extract-excel: ${err.message}\n`);
      if (err.hint) process.stderr.write(`  ${err.hint}\n`);
      process.exitCode = 1;
    } else {
      process.stderr.write(`extract-excel: unexpected error: ${(err as Error).message}\n`);
      process.exitCode = 2;
    }
  });
