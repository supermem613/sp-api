#!/usr/bin/env node
'use strict';

const { main } = require('../src/sp-api-core');

// Set process.exitCode and let the event loop drain rather than calling
// process.exit(). A forced process.exit() tears down libuv while undici's
// HTTP/2 session for SharePoint is still open, which double-closes an async
// handle and aborts the process on Windows with the async.c UV_HANDLE_CLOSING
// assertion. Every main() exit path is already followed by a return, so
// control flow does not depend on process.exit() halting execution.
main(process.argv.slice(2), {
  stdout: process.stdout,
  stderr: process.stderr,
  exit: code => { process.exitCode = code; },
}).catch(err => {
  process.stdout.write(JSON.stringify({
    ok: false,
    command: 'sp-api',
    data: null,
    error: { code: 'UNHANDLED_ERROR', message: err.message },
    meta: { schemaVersion: '0.1.0' },
  }, null, 2) + '\n');
  process.exitCode = 1;
});
