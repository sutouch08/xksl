<?php
defined('BASEPATH') OR exit('No direct script access allowed');
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Shared\Date;

class Sales_channels_details extends PS_Controller
{
  public $menu_code = 'RSOCDS';
	public $menu_group_code = 'RE';
  public $menu_sub_group_code = 'RESALE';
	public $title = 'รายงานออเดอร์ออนไลน์แสดงรายละเอียดการจัดส่ง(ไม่ใช่ยอดขาย)';
  public $filter;
  public function __construct()
  {
    parent::__construct();
    $this->home = base_url().'report/sales/sales_channels_details';
    $this->load->model('report/sales/sales_report_model');
  }

  public function index()
  {
    $this->load->model('masters/channels_model');
    $channels = $this->channels_model->get_online_list();
    $ds['channels_list'] = $channels;
    $this->load->view('report/sales/report_sales_channels_details', $ds);
  }


  public function do_export()
  {
    ini_set('memory_limit','2048M'); // This also needs to be increased in some cases. Can be changed to a higher value as per need)
    ini_set('sqlsrv.ClientBufferMaxKBSize','2097152'); // Setting to 2048M
    ini_set('sqlsrv.client_buffer_max_kb_size','2097152'); // Setting to 512M - for pdo_sqlsrv
    
    $allChannels = $this->input->post('allChannels');
    $channels = $this->input->post('channels');
    $pdFrom = $this->input->post('pdFrom');
    $pdTo = $this->input->post('pdTo');
    $refCodeFrom = $this->input->post('refCodeFrom');
    $refCodeTo = $this->input->post('refCodeTo');
    $fromDate = $this->input->post('fromDate');
    $toDate = $this->input->post('toDate');

    $allProduct = (!empty($pdFrom) && !empty($pdTo)) ? 0 : 1;

    $token = $this->input->post('token');

    $ds = array(
      'all_channels' => $this->input->post('allChannels'),
      'channels' => $this->input->post('channels'),
      'item_from' => $this->input->post('pdFrom'),
      'item_to' => $this->input->post('pdTo'),
      'from_reference' => $this->input->post('refCodeFrom'),
      'to_reference' => $this->input->post('refCodeTo'),
      'from_date' => $this->input->post('fromDate'),
      'to_date' => $this->input->post('toDate')
    );

    $result = $this->sales_report_model->get_online_channels_details($ds);


    $ch_title = "";
    if(!empty($channels))
    {
      $i = 1;
      foreach($channels as $id_channels)
      {
        $ch_title .= $i == 1 ? $id_channels :', '.$id_channels;
        $i++;
      }
    }

    //---  Report title
    $report_title = 'รายงาน ออเดอร์ออนไลน์ แสดงรายละเอียดการจัดส่ง วันที่ ' . thai_date($fromDate,'/') .' ถึง '.thai_date($toDate, '/');
    $channels_title = 'ช่องทางการขาย : '.($allChannels == 1 ? 'ทั้งหมด' : $ch_title);
    $pd_title     = 'สินค้า :  '. ($allProduct == 1 ? 'ทั้งหมด' : '('.$pdFrom.') - ('.$pdTo.')');

    //--- load excel library
    $excel = new Spreadsheet();
    $sheet = $excel->getActiveSheet();
    $sheet->setTitle('Sales online channels Report');

    //--- set report title header
    $sheet->setCellValue('A1', $report_title);
    $sheet->mergeCells('A1:G1');
    $sheet->setCellValue('A2', $channels_title);
    $sheet->mergeCells('A2:G2');
    $sheet->setCellValue('A3', $pd_title);
    $sheet->mergeCells('A3:G3');
    $sheet->setCellValue('A4', 'วันที่เอกสาร : ('.thai_date($fromDate,'/') .') - ('.thai_date($toDate,'/').')');
    $sheet->mergeCells('A4:G4');

    //--- set Table header


    $sheet->setCellValue('A5', 'ลำดับ');
    $sheet->setCellValue('B5', 'วันที่');
    $sheet->setCellValue('C5', 'เอกสาร');
    $sheet->setCellValue('D5', 'อ้างอิง');
    $sheet->setCellValue('E5', 'เลขที่จัดส่ง');
    $sheet->setCellValue('F5', 'ชื่อลูกค้า');
    $sheet->setCellValue('G5', 'ที่อยู่บรรทัด 1');
    $sheet->setCellValue('H5', 'ที่อยู่บรรทัด 2');
    $sheet->setCellValue('I5', 'อำเภอ');
    $sheet->setCellValue('J5', 'จังหวัด');
    $sheet->setCellValue('K5', 'รหัสไปรษณีย์');
    $sheet->setCellValue('L5', 'เบอร์โทรศัพท์');
    $sheet->setCellValue('M5', 'ช่องทางขาย');
    $sheet->setCellValue('N5', 'ช่องทางการชำระเงิน');
    $sheet->setCellValue('O5', 'สินค้า');
    $sheet->setCellValue('P5', 'ราคา');
    $sheet->setCellValue('Q5', 'จำนวน');
    $sheet->setCellValue('R5', 'ส่วนลด');
    $sheet->setCellValue('S5', 'มูลค่า');
    $sheet->setCellValue('T5', 'ค่าจัดส่ง');
    $sheet->setCellValue('U5', 'ค่าบริการ');
    $sheet->setCellValue('V5', 'สถานะ');
    $sheet->setCellValue('W5', 'ขนส่ง');

    //---- กำหนดความกว้างของคอลัมภ์
    $sheet->getColumnDimension('B')->setWidth(15);
    $sheet->getColumnDimension('C')->setWidth(20);
    $sheet->getColumnDimension('D')->setWidth(20);
    $sheet->getColumnDimension('E')->setWidth(20);
    $sheet->getColumnDimension('F')->setWidth(25);
    $sheet->getColumnDimension('G')->setWidth(90);
    $sheet->getColumnDimension('H')->setWidth(50);
    $sheet->getColumnDimension('I')->setWidth(20);
    $sheet->getColumnDimension('J')->setWidth(20);
    $sheet->getColumnDimension('K')->setWidth(15);
    $sheet->getColumnDimension('L')->setWidth(15);
    $sheet->getColumnDimension('M')->setWidth(15);
    $sheet->getColumnDimension('N')->setWidth(15);
    $sheet->getColumnDimension('O')->setWidth(25);
    $sheet->getColumnDimension('T')->setWidth(15);
    $sheet->getColumnDimension('U')->setWidth(15);
    $sheet->getColumnDimension('V')->setWidth(15);
    $sheet->getColumnDimension('W')->setWidth(15);

    $row = 6;


    if(!empty($result))
    {
      $no = 1;
      $prev_code = NULL;
      $this->load->model('address/address_model');
      $adr = NULL;
      foreach($result as $rs)
      {        
        $date = Date::PHPToExcelObject($rs->date_add);

        if($prev_code != $rs->code)
        {
          if(!empty($rs->id_address))
          {
            $adr = $this->address_model->get_shipping_detail($rs->id_address);
          }
          else
          {
            $adr = $this->address_model->get_shipping_address_by_code($rs->customer_ref);
          }
        }

        if(empty($adr))
        {
          $adr = new stdClass();
          $adr->name = NULL;
          $adr->address = NULL;
          $adr->sub_district = NULL;
          $adr->district = NULL;
          $adr->province = NULL;
          $adr->postcode = NULL;
          $adr->phone = NULL;
        }


        //--- ลำดับ
        $sheet->setCellValue('A'.$row, $no);

        //--- วันที่เอกสาร
        $sheet->setCellValue('B'.$row, $date);

        //--- เลขที่เอกสาร (SO)
        $sheet->setCellValue('C'.$row, $rs->code);

        //--- เลขที่อ้างอิง
        $sheet->setCellValueExplicit('D'.$row, $rs->reference, DataType::TYPE_STRING);

        //--- เลขที่จัดส่ง
        $tracking = empty($rs->tracking_no) ? $rs->shipping_code : $rs->tracking_no;
        $sheet->setCellValueExplicit('E'.$row, $tracking, DataType::TYPE_STRING);

        //--- ชือผู้รับสินค้า
        $sheet->setCellValue('F'.$row, $adr->name);

        //--- ที่อยู่บรรทัดที่ 1
        $sheet->setCellValue('G'.$row, $adr->address);

        //--- ที่อยู่บรรทัดที่ 2
        $sheet->setCellValue('H'.$row, $adr->sub_district);

        //--- อำเภอ / เขต
        $sheet->setCellValue('I'.$row, $adr->district);

        //--- จังหวัด
        $sheet->setCellValue('J'.$row, $adr->province);

        //--- รหัรหัสไปรษณีย์
        $sheet->setCellValueExplicit('K'.$row, $adr->postcode, DataType::TYPE_STRING);

        //--- เบอร์โทรศัพท์
        $sheet->setCellValueExplicit('L'.$row, $adr->phone, DataType::TYPE_STRING);
        //--- ช่องทางการขาย
        $sheet->setCellValue('M'.$row, $rs->channels);

        //--- ช่องทางการชำระเงิน
        $sheet->setCellValue('N'.$row, $rs->payment);

        //--- รหัสสินค้า
        $sheet->setCellValue('O'.$row, $rs->product_code);

        //--- ราคาสินค้า
        $sheet->setCellValue('P'.$row, $rs->price);

        //--- จำนวน
        $sheet->setCellValue('Q'.$row, $rs->qty);

        //--- ส่วนลดรวมเป้นจำนวนเงิน
        $sheet->setCellValue('R'.$row, $rs->discount_amount);

        //--- ยอดเงินรวม
        $sheet->setCellValue('S'.$row, $rs->total_amount);

        //--- ค่าจัดส่ง
        $sheet->setCellValue('T'.$row, $rs->shipping_fee);

        //--- ค่าบริการ
        $sheet->setCellValue('U'.$row, $rs->service_fee);

        //--- สถานะออเดอร์
        $sheet->setCellValue('V'.$row, $rs->state);

        //--- ขนส่ง
        $sheet  ->setCellValue('W'.$row, $rs->sender);

        $no++;
        $row++;
      }

      $sheet->getStyle('B6:B'.$row)->getNumberFormat()->setFormatCode('dd/mm/yyyy');
    }


    setToken($token);
    $file_name = "Repor Sale Online Channels Details.xlsx";
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); /// form excel 2007 XLSX
    header('Content-Disposition: attachment;filename="'.$file_name.'"');
    $writer = new Xlsx($excel);
    $writer->save('php://output');
  }
}
