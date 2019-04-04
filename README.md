# importAndExport
依赖于phpoffice/phpspreadsheet实现的通用导入、导出功能。

ImportExecl()
/**
     * @param string $pfilename
     *               文件地址
     * @param string $title 
     *            execl标题列，如：姓名、性别、地址
     * @param int $sheet
     *            文件行数 默认从第一行取
     * @param int $columnCnt
     *            文件列  默认最大列
     * @throws \Exception
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     *
     */

exportExecl
  /**
     * @param array $title 例如：[姓名,性别]
     * @param array $data  例如：[['nickname'=>'张三','sex'=>'男'],
     *                           ['nickname'=>'李四','sex'=>'女']]
     * @param string $fileName
     * @return bool
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     *
     */
