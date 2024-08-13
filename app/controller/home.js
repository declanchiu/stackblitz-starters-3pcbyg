const { Controller } = require('egg');
const fs = require('fs');
const ExcelJS = require('exceljs');

class HomeController extends Controller {
  async index() {
    async function readImageFromExcel(filePath, cellAddress) {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1);

      const cell = worksheet.getCell(cellAddress);

      let resbonse = null;
      worksheet.getImages().forEach((image) => {
        const { range } = image;
        if (
          range.tl &&
          range.tl.row === cell.row - 1 &&
          range.tl.col === cell.col - 1
        ) {
          const imageId = image.imageId;
          const workbookImage = workbook.model.media.find((img) => {
            console.log(img.index, imageId);
            return img.index === imageId;
          });
          if (workbookImage) {
            const buffer = Buffer.from(workbookImage.buffer);
            resbonse = buffer;
          }
        }
      });

      return resbonse;
    }

    const filePath = `${__dirname}/demo.xlsx`;
    const cellAddress = 'A2';
    const res = await readImageFromExcel(filePath, cellAddress);
    console.log('读取结果', res);

    await this.ctx.render('home.tpl', { name: 'egg' });
  }
}

module.exports = HomeController;
