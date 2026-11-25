<?php
defined('BASEPATH') OR exit('No direct script access allowed');
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Warehouse extends PS_Controller
{
  public $menu_code = 'DBWRHS';
	public $menu_group_code = 'DB';
  public $menu_sub_group_code = 'WAREHOUSE';
	public $title = 'เพิ่ม/แก้ไข คลังสินค้า';
  public $segment = 4;

  public function __construct()
  {
    parent::__construct();
    $this->home = base_url().'masters/warehouse';
    $this->load->model('masters/warehouse_model');
    $this->load->helper('warehouse');
  }

  public function index()
  {
    $filter = array(
      'code' => get_filter('code', 'wh_code', ''),
      'role' => get_filter('role', 'wh_role', 'all'),
      'is_consignment' => get_filter('is_consignment', 'is_consignment', 'all'),
      'active' => get_filter('active', 'wh_active', 'all'),
      'sell' => get_filter('sell', 'wh_sell', 'all'),
      'prepare' => get_filter('prepare', 'wh_prepare', 'all'),
      'lend' => get_filter('lend', 'wh_lend', 'all'),
      'auz' => get_filter('auz', 'wh_auz', 'all'),
      'is_pos' => get_filter('is_pos', 'wh_is_pos', 'all')
    );
		
		$perpage = get_rows();
		$rows = $this->warehouse_model->count_rows($filter);
		$list = $this->warehouse_model->get_list($filter, $perpage, $this->uri->segment($this->segment));

    if(!empty($list))
    {
      foreach($list as $rs)
      {
        $rs->zone_count = $this->warehouse_model->count_zone($rs->code);
      }
    }

    $filter['list'] = $list;
		$init = pagination_config($this->home.'/index/', $rows, $perpage, $this->segment);
		$this->pagination->initialize($init);
    $this->load->view('masters/warehouse/warehouse_list', $filter);
  }


  public function edit($code)
  {
    if($this->pm->can_edit)
    {
      $ds['ds'] = $this->warehouse_model->get($code);
      $this->load->view('masters/warehouse/warehouse_edit', $ds);
    }
    else
    {
      set_error("คุณไม่มีสิทธิ์แก้ไขคลังสินค้า");
      redirect($this->home);
    }
  }


  public function update()
  {
    if($this->pm->can_edit)
    {
      if($this->input->post('code'))
      {
        $code = $this->input->post('code');        

        $arr = array(
          'role' => $this->input->post('role'),
          'sell' => $this->input->post('sell'),
          'prepare' => $this->input->post('prepare'),
          'lend' => $this->input->post('lend'),
          'auz' => $this->input->post('auz'),
          'active' => $this->input->post('active'),
          'is_consignment' => get_null($this->input->post('is_consignment')),
          'is_pos' => $this->input->post('is_pos') == 1 ? 1 : 0,
          'update_user' => get_cookie('uname')
        );

        if($this->warehouse_model->update($code, $arr))
        {
          set_message("Update Successfull");
          redirect($this->home.'/edit/'.$code);
        }
        else
        {
          set_error("Update Fail");
          redirect($this->home.'/edit/'.$code);
        }
      }
      else
      {
        set_error('ไม่พบรหัสคลังสินค้า');
        redirect($this->home);
      }
    }
    else
    {
      set_error('คุณไม่มีสิทธิ์แก้ไขคลังสินค้า');
      redirect($this->home);
    }
  }


  public function delete($code)
  {
    $sc = TRUE;

    if($this->pm->can_delete)
    {
      //---- count member if exists reject action
      if($this->warehouse_model->has_zone($code))
      {
        $sc = FALSE;
        $this->error = 'ไม่สามารถลบคลังได้เนื่องจากยังมีโซนอยู่';
      }
      //--- check warehouse in SAP if exists reject action
      else if($this->warehouse_model->is_sap_exists($code))
      {
        $sc = FALSE;
        $this->error = 'ไม่สามารถลบคลังได้เนื่องจากยังไม่ได้ลบคลังใน SAP';
      }
      else
      {
        if($this->warehouse_model->delete($code) === FALSE)
        {
          $sc = FALSE;
          $this->error = 'ลบคลังไม่สำเร็จ';
        }
      }
    }
    else
    {
      $sc = FALSE;
      $this->error = 'คุณไม่มีสิทธิ์ลบคลังสินค้า';
    }

    echo $sc === TRUE ? 'success' : $this->error;
  }


  public function syncData()
  {
    $last_sync = $this->warehouse_model->get_last_sync_date();    
    $newData = $this->warehouse_model->get_new_data($last_sync);

    if(!empty($newData))
    {
      foreach($newData as $rs)
      {
        if($this->warehouse_model->is_exists($rs->code))
        {
          $ds = array(
            'name' => $rs->name,
            'active' => $rs->Inactive == 'Y' ? 0 : 1,
            'last_sync' => date('Y-m-d H:i:s'),
            'update_user' => 'SAP',
            'old_code' => $rs->old_code,
            'limit_amount' => $rs->limit_amount
          );

          $this->warehouse_model->update($rs->code, $ds);
        }
        else
        {
          $ds = array(
            'code' => $rs->code,
            'name' => $rs->name,
            'active' => $rs->Inactive == 'Y' ? 0 : 1,
            'last_sync' => date('Y-m-d H:i:s'),
            'update_user' => 'SAP',
            'old_code' => $rs->old_code,
            'limit_amount' => $rs->limit_amount
          );

          $this->warehouse_model->add($ds);
        }
      }
    }

    echo 'done';
  }


  public function syncAllData()
  {
    $last_sync = date('Y-m-d H:i:s', strtotime('2019-01-01 00:00:00'));
    $newData = $this->warehouse_model->get_new_data($last_sync);

    if(!empty($newData))
    {
      foreach($newData as $rs)
      {
        if($this->warehouse_model->is_exists($rs->code))
        {
          $ds = array(
            'name' => $rs->name,
            'active' => $rs->Inactive == 'Y' ? 0 : 1,
            'last_sync' => date('Y-m-d H:i:s'),
            'update_user' => 'SAP',
            'old_code' => $rs->old_code,
            'limit_amount' => $rs->limit_amount
          );

          $this->warehouse_model->update($rs->code, $ds);
        }
        else
        {
          $ds = array(
            'code' => $rs->code,
            'name' => $rs->name,
            'active' => $rs->Inactive == 'Y' ? 0 : 1,
            'last_sync' => date('Y-m-d H:i:s'),
            'update_user' => 'SAP',
            'old_code' => $rs->old_code,
            'limit_amount' => $rs->limit_amount
          );

          $this->warehouse_model->add($ds);
        }
      }
    }

    echo 'done';
  }


  public function export_filter()
  {
    $filter = array(
      'code' => get_filter('whCode', 'wh_code', ''),
      'role' => get_filter('whRole', 'wh_role', 'all'),
      'is_consignment' => get_filter('whIsConsignment', 'is_consignment', 'all'),
      'active' => get_filter('whActive', 'wh_active', 'all'),
      'sell' => get_filter('whSell', 'wh_sell', 'all'),
      'lend' => get_filter('whLend', 'wh_lend', 'all'),
      'prepare' => get_filter('whPrepare', 'wh_prepare', 'all'),
      'auz' => get_filter('whAuz', 'wh_auz', 'all'),
      'is_pos' => get_filter('whIsPos', 'wh_is_pos', 'all')
    );

    $token = $this->input->post('token');

    $list = $this->warehouse_model->get_list($filter);

    $excel = new Spreadsheet();
    $sheet = $excel->getActiveSheet();

    $excel->setActiveSheetIndex(0);
    $sheet->setTitle('Zone master data');

    //--- set Table header


    $sheet->setCellValue('A1', 'No.');
    $sheet->setCellValue('B1', 'Code');
    $sheet->setCellValue('C1', 'Description');
    $sheet->setCellValue('D1', 'Role');
    $sheet->setCellValue('E1', 'Bin Location');
    $sheet->setCellValue('F1', 'Sell');
    $sheet->setCellValue('G1', 'Pick');
    $sheet->setCellValue('H1', 'Can be negative');
    $sheet->setCellValue('I1', 'Active');
    $sheet->setCellValue('J1', 'Is Consignment');
    $sheet->setCellValue('K1', 'Limit Amount');


    //---- กำหนดความกว้างของคอลัมภ์
    $sheet->getColumnDimension('B')->setWidth(20);
    $sheet->getColumnDimension('C')->setWidth(40);
    $sheet->getColumnDimension('D')->setWidth(20);
    $sheet->getColumnDimension('E')->setWidth(20);
    $sheet->getColumnDimension('F')->setWidth(20);
    $sheet->getColumnDimension('G')->setWidth(20);
    $sheet->getColumnDimension('H')->setWidth(20);
    $sheet->getColumnDimension('I')->setWidth(20);
    $sheet->getColumnDimension('J')->setWidth(20);
    $sheet->getColumnDimension('K')->setWidth(20);


    $row = 2;


    if(! empty($list))
    {
      $no = 1;

      foreach($list as $rs)
      {
        $sheet->setCellValue('A'.$row, $no);
        $sheet->setCellValue('B'.$row, $rs->code);
        $sheet->setCellValue('C'.$row, $rs->name);
        $sheet->setCellValue('D'.$row, $rs->role_name);
        $sheet->setCellValue('E'.$row, $this->warehouse_model->count_zone($rs->code));
        $sheet->setCellValue('F'.$row, ($rs->sell ? 'Y' : 'N'));
        $sheet->setCellValue('G'.$row, ($rs->prepare ? 'Y' : 'N'));
        $sheet->setCellValue('H'.$row, ($rs->auz ? 'Y' : 'N'));
        $sheet->setCellValue('I'.$row, ($rs->active ? 'Y' : 'N'));
        $sheet->setCellValue('J'.$row, ($rs->is_consignment ? 'Y' : 'N'));
        $sheet->setCellValue('K'.$row, $rs->limit_amount);
        $no++;
        $row++;
      }

      setToken($token);
      $file_name = "Warehouse Master Data.xlsx";
      header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); /// form excel 2007 XLSX
      header('Content-Disposition: attachment;filename="'.$file_name.'"');
      $writer = new Xlsx($excel);
      $writer->save('php://output');
    }
  }


  public function clear_filter()
  {
    $filter = array('wh_code', 'wh_role', 'is_consignment', 'wh_active', 'wh_sell', 'wh_prepare', 'wh_auz', 'wh_lend', 'wh_is_pos');
    clear_filter($filter);
  }

} 
