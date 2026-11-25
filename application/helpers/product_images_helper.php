<?php
function get_image_path($id, $size = 'default')
{
  $ci =& get_instance();
  $ci->load->model('masters/product_image_model');
  $code = $ci->product_image_model->get_style_code($id);
  $path = $ci->config->item('image_path').'products';
  $no_image_path = base_url().$path.'/no_image_'.$size.'.jpg';

  if($code !== FALSE)
  {
    $image_path = base_url().$path.'/'.$code.'/product_'.$size.'_'.$id.'.jpg';
    $file = $ci->config->item('image_file_path').'products/'.$code.'/product_'.$size.'_'.$id.'.jpg';

    if(!file_exists($file) && $size == 'large')
    {
      $file = $ci->config->item('image_file_path').'products/'.$code.'/product_large_'.$id.'.jpg';
      $image_path = base_url().$path.'/'.$code.'/product_large_'.$id.'.jpg';
      if(!file_exists($file))
      {
        $file = $ci->config->item('image_file_path').'products/'.$code.'/product_lage_'.$id.'.jpg';
        $image_path = base_url().$path.'/'.$code.'/product_lage_'.$id.'.jpg';
      }
    }

    return file_exists($file) ? $image_path : $no_image_path;
  }

  return $no_image_path;
}

function get_product_image($code, $size = 'default')
{
  $ci =& get_instance();
  $ci->load->model('masters/product_image_model');
  $id_image = $ci->product_image_model->get_id_image($code);
  return get_image_path($id_image, $size);
}

function delete_product_image($id, $code)
{
  $ci =& get_instance();
  $path = $ci->config->item('image_file_path').'products/';
  $use_size = array('mini', 'default', 'medium', 'lage', 'large');
  foreach($use_size as $size)
  {
    $image_path = $path.'/'.$code.'/product_'.$size.'_'.$id.'.jpg';
    unlink($image_path);
  }

  return TRUE;
}

function get_cover_image($code, $size = 'default')
{
  $ci =& get_instance();
  $ci->load->model('masters/product_image_model');
  $id  = $ci->product_image_model->get_cover($code);
  return get_image_path($id, $size);
}

function no_image_path($size)
{
  $ci =& get_instance();
  $path = $ci->config->item('image_path');
  $no_image_path = base_url().$path.'/no_image_'.$size.'.jpg';
  return $no_image_path;
}
