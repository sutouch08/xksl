<?php
function select_color_group($id = '')
{
  $ci =& get_instance();
  $ci->load->model('masters/product_color_model');
  $result = $ci->product_color_model->get_color_group();
  $ds = '';

  if(!empty($result))
  {
    foreach($result as $rs)
    {
      $ds .= '<option value="'.$rs->id.'" '.is_selected($rs->id, $id).'>'.$rs->code.' : '.$rs->name.'</option>';
    }
  }

  return $ds;
}
