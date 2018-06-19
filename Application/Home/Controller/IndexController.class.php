<?php

namespace Home\Controller;

use Think\Controller;
use Think\Upload;

class IndexController extends Controller
{
    public function index()
    {
        $this->show('<style type="text/css">*{ padding: 0; margin: 0; } div{ padding: 4px 48px;} body{ background: #fff; font-family: "微软雅黑"; color: #333;font-size:24px} h1{ font-size: 100px; font-weight: normal; margin-bottom: 12px; } p{ line-height: 1.8em; font-size: 36px } a,a:hover{color:blue;}</style><div style="padding: 24px 48px;"> <h1>:)</h1><p>欢迎使用 <b>ThinkPHP</b>！</p><br/>版本 V{$Think.version}</div><script type="text/javascript" src="http://ad.topthink.com/Public/static/client.js"></script><thinkad id="ad_55e75dfae343f5a1"></thinkad><script type="text/javascript" src="http://tajs.qq.com/stats?sId=9347272" charset="UTF-8"></script>', 'utf-8');
    }

    public function Excelshow()
    {
        $this->display('excelhello');

    }


    //Excel表格导出到本地
    public function data1()
    {
        import("Org.Util.PHPExcel.PHPExcel");
        //实例化PHPExcel类，等同于在桌面上新建一个excel
        $objPHPExcel = new \PHPExcel();
        $objSheet = $objPHPExcel->getActiveSheet(0);
        $objSheet->setTitle('商品详情');//设置sheet标题
        $model = M('data1');
        $data = $model->select(); //获取当前表中的所有数据
        // 填充数据
        $objSheet->setCellValue('A1', '序号')
            ->setCellValue('B1', '商品名')
            ->setCellValue('C1', '价格')
            ->setCellValue('D1', '数量');
        $j = 2;
        foreach ($data as $key => $val) {
            $objSheet->setCellValue("A" . $j, $val['pid'])
                ->setCellValue("A" . $j, $val['pid'])
                ->setCellValue("B" . $j, $val['pname'])
                ->setCellValue("C" . $j, $val['pprice'])
                ->setCellValue("D" . $j, $val['pcount']);
            $j++;
        }
        $objwriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        // $objwriter->save('export_1.xls');//保存文件
        $this->export_browser('Excel5', 'broswer.xls');
        $objwriter->save("php://output");
    }
    //Excel表格输出到浏览器
    protected function export_browser($type, $filename)
    {
        if ($type == "Excel5") {
            //告诉浏览器将要输出Excel03文件
            header('Content-Type: application/vnd.ms-excel');
        } else {
            // 告诉浏览器将要输出excel07文件
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        }
        // 告诉浏览器将输出文件的名称
        header('Content-Disposition: attachment;filename="' . $filename . '"');
        // 禁止缓存
        header('Cache-Control: max-age=0');

    }

    //Excel文件导出形成图表样式
    public function excel_chart()
    {
        import("Org.Util.PHPExcel.PHPExcel");
        //实例化PHPExcel类，等同于在桌面上新建一个excel
        $objPHPExcel = new \PHPExcel();
        $objSheet = $objPHPExcel->getActiveSheet(0);
        $objSheet->setTitle('商品详情');//设置sheet标题
        $model = M('data1');
        $data = $model->select(); //获取当前表中的所有数据
        // 填充数据
        $objSheet->setCellValue('A1', '序号')
            ->setCellValue('B1', '商品名')
            ->setCellValue('C1', '价格')
            ->setCellValue('D1', '数量');
        $j = 2;
        foreach ($data as $key => $val) {
            $objSheet->setCellValue("A" . $j, $val['pid'])
                ->setCellValue("A" . $j, $val['pid'])
                ->setCellValue("B" . $j, $val['pname'])
                ->setCellValue("C" . $j, $val['pprice'])
                ->setCellValue("D" . $j, $val['pcount']);
            $j++;
        }

        // 开始图表代码的编写
        // 先取得绘制图表的标签
        $labels = array(
            new \PHPExcel_Chart_DataSeriesValues("String", '商品详情!$C$1', null, 1),
            new \PHPExcel_Chart_DataSeriesValues("String", '商品详情!$D$1', null, 1),
        );
        // 取得图表X轴的刻度
        $xLabels = array(
            new \PHPExcel_Chart_DataSeriesValues("String", '商品详情!$B$2:$B$4', null, 3),
        );
        // 取得绘图所需的数据
        $datas = array(
            new \PHPExcel_Chart_DataSeriesValues("String", '商品详情!$C$2:$C$4', null, 3),
        );

        // 根据取得的东西做出一个图表的框架
        $series = array(
            new \PHPExcel_Chart_DataSeries(
                \PHPExcel_Chart_DataSeries::TYPE_LINECHART,
                \PHPExcel_Chart_DataSeries::GROUPING_STANDARD,
                range(0, count($labels) - 1),
                $labels,
                $xLabels,
                $datas
            )
        );
        $layout = new \PHPExcel_Chart_Layout();
        $layout->setShowVal(true);
        $areas = new \PHPExcel_Chart_PlotArea($layout, $series);
        $legend = new \PHPExcel_Chart_Legend(\PHPExcel_Chart_Legend::POSITION_RIGHT, $layout, false);
        $title = new \PHPExcel_Chart_Title("商品详情");
        $ytitle = new \PHPExcel_Chart_Title("value(价格)");
        //生成一个图表
        $chart = new \PHPExcel_Chart(
            "line_chart", $title, $legend, $areas, true, false, null, $ytitle
        );
        // 给定图表所在的表格中的位置
        $chart->setTopLeftPosition('A7')->setBottomRightPosition('K25');
        // 将chart添加到表格中
        $objSheet->addChart($chart);


        // 生成excel文件
        $objwriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objwriter->setIncludeCharts(true);
        // $objwriter->save('export_1.xls');//保存文件
        $this->export_browser('Excel2007', 'broswer.xls');
        $objwriter->save("php://output");
    }


    //Excel文件导入功能实现
    public function upload()
    {
        ini_set('memory_limit', '1024M');
        if (!empty($_FILES)) {
            $config = array(
                'maxSize' => 3145728000,
                'rootPath' => "./Public/",
                'savePath' => 'Uploads/',
                'subName' => array('date', 'Ymd'),
            );
            $upload = new \Think\Upload($config);
            if (!$info = $upload->upload()) {
                $this->error($upload->getError());
            }
            import("Org.Util.PHPExcel.PHPExcel");
            import("Org.Util.PHPExcel..PHPExcel.Reader.Excel5");
            $file_name = $upload->rootPath . $info['excel']['savepath'] . $info['excel']['savename'];
            // $extension = strtolower(pathinfo($file_name, PATHINFO_EXTENSION));//判断导入表格后缀格式
            // if ($extension == 'xlsx') {
            //     $objReader = \PHPExcel_IOFactory::createReader('Excel2007');
            //     $objPHPExcel = $objReader->load($file_name, $encode = 'utf-8');
            // } else if ($extension == 'xls') {
            //     $objReader = \PHPExcel_IOFactory::createReader('Excel5');
            //     $objPHPExcel = $objReader->load($file_name, $encode = 'utf-8');
            // }
            $filetype = \PHPExcel_IOFactory::identify($file_name);
            $objReader = \PHPExcel_IOFactory::createReader($filetype);
            $objPHPExcel = $objReader->load($file_name, $encode = 'utf-8');
            $sheet = $objPHPExcel->getActiveSheet(0);
            $highestRow = $sheet->getHighestRow();//取得总行数
            $highestColumn = $sheet->getHighestColumn(); //取得总列数
            D('data1')->execute('truncate table data1');
            for ($i = 2; $i <= $highestRow; $i++) {
                //看这里看这里,前面小写的a是表中的字段名，后面的大写A是excel中位置
                $data['pid'] = $sheet->getCell("A" . $i)->getValue();
                $data['pname'] = $sheet->getCell("B" . $i)->getValue();
                $data['pprice'] = $sheet->getCell("C" . $i)->getValue();
                $data['pcount'] = $sheet->getCell("D" . $i)->getValue();
                //看这里看这里,这个位置写数据库中的表名

                D('data1')->add($data);
            }
            $this->success('导入成功!');
        } else {
            $this->error("请选择上传的文件");
        }
    }


}