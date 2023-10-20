import { message } from 'antd';
import saveAs from 'file-saver';
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore
import ExportExcelWorker from './export.worker?worker';
import { ExcelParamsType, exportExcel } from './export.worker';

// 导出函数，传入参数props，类型为ExcelParamsType，根据浏览器兼容性判断是否启用worker
export function createExportWorker(props: ExcelParamsType) {
  const { isWorker = false } = props;

  // 如果当前浏览器不支持webworker，或者没有手动指定worker，则使用原来的方法进行导出
  if (!window.Worker || !isWorker) {
    // 如果当前浏览器不支持webworker，则使用原来的方法进行导出
    exportExcel({ ...props, isWorker: false });
    return;
  }
  // 创建一个ExportExcelWorker实例
  const worker = new ExportExcelWorker();

  // 向webworker发送消息，开始导出
  worker.postMessage({
    type: 'start',
    data: props
  });

  // 监听webworker的消息
  worker.onmessage = function (e) {
    const { data, type } = e.data || {};

    // 成功则下载，并且终止worker
    if (type === 'success') {
      saveAs(data, props.fileName);
      worker.terminate();
    }
    // 失败message提示
    if (type === 'error') {
      message.error(data);
    }
  };
}
