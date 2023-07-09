import { createWriteStream, rmdir, unlink } from 'fs';
import archiver from 'archiver';
import { promisify } from 'util';

const rmdirD = promisify(rmdir);

export async function makeZip(id) {
  var output = createWriteStream('data/zipData/'+id+'.zip');
  var archive = archiver('zip');

  output.on('close', function () {
      console.log(archive.pointer() + ' total bytes');
      console.log('archiver has been finalized and the output file descriptor has closed.');
  });

  archive.on('error', function(err){
      throw err;
  });

  archive.pipe(output);

  // append files from a sub-directory, putting its contents at the root of archive
  archive.directory('data/pdfData/'+id, false);

  await archive.finalize();
  try {
    await rmdirD('data/pdfData/' + id, { recursive: true });
    return true;
  } catch (err) {
    console.error(err);
    return false;
  }
}