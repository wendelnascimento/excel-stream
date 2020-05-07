#!/usr/bin/env node

const fs = require('fs');
const os = require('os');
const path = require('path');
const chpro = require('child_process');

const csv = require('csv-stream');
const osenv = require('osenv');
const concat = require('concat-stream');
const eventStream = require('event-stream');

let { spawn } = chpro;
if (os.type() === 'Windows_NT') spawn = require('cross-spawn');

module.exports = function (options) {
  const read = eventStream.through();
  let duplex;

  const filename = path.join(osenv.tmpdir(), `_${Date.now()}`);

  const spawnArgs = [];

  if (options) {
    options.sheet && spawnArgs.push('--sheet') && spawnArgs.push(options.sheet) && delete options.sheet;
    options.sheetIndex && spawnArgs.push('--sheet-index') && spawnArgs.push(options.sheetIndex) && delete options.sheetIndex;
  }

  spawnArgs.push(filename);

  const _data = {};

  function getSheet(childSpawnArgs) {
    const child = spawn(require.resolve('xlsx/bin/xlsx.njs'), childSpawnArgs);

    child.on('exit', (code) => {
      if (code === null || code !== 0) {
        child.stderr.pipe(concat((errstr) => {
          duplex.emit('error', new Error(errstr));
        }));
      }
    });
    return child.stdout.pipe(csv.createStream(options));
  }

  const write = fs.createWriteStream(filename)
    .on('close', () => {
      if (options.allSheets) {
        const streams = [];
        const childCount = spawn(require.resolve('xlsx/bin/xlsx.njs'), ['--list-sheets', filename]);
        childCount.stdout
          .pipe(eventStream.split())
          .pipe(eventStream.filterSync((sheetName) => sheetName))
          .pipe(eventStream.mapSync((sheetName) => getSheet(['--sheet', sheetName, filename])))
          .pipe(eventStream.mapSync((item) => {
            streams.push(item);
            return item;
          }))
          .on('end', () => {
            eventStream.merge(streams)
              .pipe(eventStream.through(function (data) {
                const keys = Object.keys(data);
                keys.forEach((k) => {
                  const value = data[k].trim();
                  _data[k.trim()] = isNaN(value) ? value : +value;
                });
                this.queue(_data);
              }))
              .pipe(read);
          });
      } else {
        getSheet(spawnArgs);
      }
    });

  return (duplex = eventStream.duplex(write, read));
};


if (!module.parent) {
  const JSONStream = require('JSONStream');
  const args = require('minimist')(process.argv.slice(2));
  process.stdin
    .pipe(module.exports())
    .pipe(args.lines || args.newlines
      ? JSONStream.stringify('', '\n', '\n', 0)
      : JSONStream.stringify())
    .pipe(process.stdout);
}
