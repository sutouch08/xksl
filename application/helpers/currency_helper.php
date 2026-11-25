<?php

function select_currency($code = '')
{
  $sc = '';
  $ci =& get_instance();
  $ci->load->model('tool_model');
  $options = $ci->tool_model->get_currency();

  if(!empty($options))
  {
    $df = getConfig('CURRENCY');
		
    foreach($options as $rs)
    {
      if(empty($code))
      {
        $sc .= '<option value="'.$rs->code.'" '.is_selected($df, $rs->code).'>'.$rs->code.'</option>';
      }
      else
      {
        $sc .= '<option value="'.$rs->code.'" '.is_selected($code, $rs->code).'>'.$rs->code.'</option>';
      }
    }
  }

  return $sc;
}

