<?php
function select_active_collection($code = '')
{
  $ci =& get_instance();
  $ci->load->model('masters/product_collection_model');  
  $result = $ci->product_collection_model->get_all();
  $ds = '';

  if(!empty($result))
  {
    foreach($result as $rs)
    {
      if($rs->active == 1 OR $rs->code == $code)
      {
        $ds .= '<option value="'.$rs->code.'" '.is_selected($rs->code, $code).'>'.$rs->name.'</option>';        
      }
    }
  }

  return $ds;
}


function select_all_collection($code = '')
{
  $ci =& get_instance();
  $ci->load->model('masters/product_collection_model');
  $result = $ci->product_collection_model->get_all();
  $ds = '';
  
  if(!empty($result))
  {
    foreach($result as $rs)
    {
      $ds .= '<option value="'.$rs->code.'" '.is_selected($rs->code, $code).'>'.$rs->name.'</option>';
    }
  }

  return $ds;
}
