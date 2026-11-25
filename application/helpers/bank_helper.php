<?php
function bankLogoUrl($code)
{
  $ci =& get_instance();
  $img  = $code.'.png';
  $path	= base_url().$ci->config->item('image_path').'banks/';
  $image_path = $path.$img;
  $noimg = $path.'noimg.png';
 	$file = $ci->config->item('image_file_path').'banks/'.$img;

 	if( ! file_exists($file) )
 	{
 		return $noimg;
 	}

 	return $image_path;
}


function select_bank_account($id = '')
{
  $sc = '';
  $ci =& get_instance();
  $ci->load->model('masters/bank_model');
  $banks = $ci->bank_model->get_data();
  if(!empty($banks))
  {
    foreach($banks as $rs)
    {
      $sc .= '<option value="'.$rs->id.'" '.is_selected($rs->id, $id).'>'.$rs->bank_name.' : '.$rs->acc_no.'</option>';
    }
  }

  return $sc;
}


function select_bank($code = '')
{
  $sc = "";
  $ci =& get_instance();
  $ci->load->model('masters/bank_model');
  $bank = $ci->bank_model->get_banks();

  if(!empty($bank))
  {
    foreach($bank as $rs)
    {
      $selected = is_selected($rs->code, $code);

      $sc .= "<option value='{$rs->code}' {$selected}>{$rs->name}</option>";
    }
  }

  return $sc;
}
