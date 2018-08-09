'use strict';

const Controller = require('egg').Controller;
const officegen = require('officegen');
const { PassThrough } = require('stream');

const fs = require('fs');
const async = require ( 'async' );


var archiver = require('archiver');




//思路
//1、生成多个文件  officegen  async
//2、将文件放在指定文件夹
//3、将问文件夹压缩  archiver
//4、将压缩的压缩包发给前端，可以用流输出给前端，也可以使用重定向通过静态文件下载的方式给前端。

class HomeController extends Controller {
  async index() {


    // 创建生成的压缩包路径，可以把它放在静态目录中
    const output = fs.createWriteStream(__dirname + '/example.zip');
    const archive = archiver('zip');


    const docx = officegen ( 'docx' );
    const pObj = docx.createP ();
    pObj.addText ( 'Hello' );

    const docx2 = officegen ( 'docx' );
    const pObj2 = docx2.createP ();
    pObj2.addText ( 'world' );


    const out = new PassThrough();
    const out4 = new PassThrough();

    var out1 = fs.createWriteStream ( 'app/out.docx' );
    var out2 = fs.createWriteStream ( 'app/out2.docx' );

    out.on ( 'error', function ( err ) {
      console.log ( err );
    });

   /* async.parallel ([
      function ( done ) {
        out.on ( 'close', function () {
          console.log ( 'Finish to create a DOCX file.' );
          done ( null );
        });
        docx.generate ( out1 );
      }

    ], function ( err ) {
      if ( err ) {
        console.log ( 'error: ' + err );
      } // Endif.
    });*/

    //生成两个文件,并保存在服务器
/*
    async.parallel({
      one: function ( done ) {
        out.on ( 'close', function () {
          console.log ( 'Finish to create a DOCX file.' );
          done ( null );
        });
        docx.generate ( out1 );
      },
      two: function ( done ) {
        out.on ( 'close', function () {
          console.log ( 'Finish to create a DOCX file.' );
          done ( null );
        });
        docx2.generate ( out2 );
      }
    }, function(err, results) {
      console.log(error);
    });
*/



   /* var file1 = __dirname + '/file1.txt';
    archive.append(fs.createReadStream(file1), { name: 'file1.txt' });*/


   // archive.pipe(output);

    var file1 = __dirname+'/out.docx';
    archive.append(fs.createReadStream(file1), {
      name: 'out.docx'
    });


    archive.finalize();

    //输出文件到前端
    docx.generate ( out );
    docx2.generate ( out4 );

    //set the archive name
    //res.attachment('archive-name.zip');

    //this is the streaming magic
    //archive.pipe(res);


    //console.log(out);
    this.ctx.type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    this.ctx.set('Content-disposition', `attachment; archive-name.zip`);
    this.ctx.body = out;
  }
}

module.exports = HomeController;
