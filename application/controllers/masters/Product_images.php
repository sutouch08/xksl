<?php
defined('BASEPATH') or exit('No direct script access allowed');
class Product_images extends PS_Controller
{
  public $menu_code = 'DBPROD';

  public function __construct()
  {
    parent::__construct();
    $this->load->model('masters/product_image_model');
    $this->load->helper('product_images');
  }

  public function upload_images($code)
  {
    if (! isset($_FILES['file']))
    {
      echo 'No file selected';
      exit;
    }

    $files = $_FILES['file'];
    $count = count($_FILES['file']['name']);
    $this->load->library('upload');
    $image_path = $this->config->item('image_path') . 'products/';
    $imagePath = $image_path . '/' . $code . '/';
    $use_size = array('mini', 'default', 'medium', 'large');

    if (! is_dir($imagePath)) mkdir($imagePath, 0755, TRUE);

    for ($i = 0; $i < $count; $i++)
    {
      $_FILES['file']['name'] = $files['name'][$i];
      $_FILES['file']['type'] = $files['type'][$i];
      $_FILES['file']['tmp_name'] = $files['tmp_name'][$i];
      $_FILES['file']['error'] = $files['error'][$i];
      $_FILES['file']['size'] = $files['size'][$i];

      $config = [
        'upload_path' => $imagePath,
        'allowed_types' => 'jpg|jpeg|png|gif|webp',
        'max_size' => 10240,
        'encrypt_name'  => FALSE,
        'overwrite' => TRUE,
        'file_name' => uniqid('temp_')
      ];

      $this->upload->initialize($config);

      if (! $this->upload->do_upload('file'))
      {
        echo $this->upload->display_errors();
        exit;
      }

      $upload = $this->upload->data();
      $source_file = $upload['full_path'];
      $id_image = $this->product_image_model->get_new_id();
      $img_name = $id_image;

      foreach ($use_size as $size)
      {
        $prop = $this->getImageSizeProperties($size);
        $target_file = $imagePath . $prop['prefix'] . $img_name . '.jpg';

        switch (strtolower($upload['image_type']))
        {
          case 'jpeg':
          case 'jpg':
            $src = imagecreatefromjpeg($source_file);
            break;

          case 'png':
            $src = imagecreatefrompng($source_file);
            break;

          case 'gif':
            $src = imagecreatefromgif($source_file);
            break;

          case 'webp':
            $src = imagecreatefromwebp($source_file);
            break;
          default:
            @unlink($source_file);
            echo 'Unsupported image type';
            exit;
        }

        if (! $src)
        {
          @unlink($source_file);
          echo 'Failed to create image resource';
          exit;
        }

        // ขนาดเดิม
        $src_width  = imagesx($src);
        $src_height = imagesy($src);
        $new_size = intval($prop['size']);

        $ratio = min(
          $new_size / $src_width,
          $new_size / $src_height
        );
        // Resize ให้กว้าง 500px และรักษาสัดส่วน
        $new_width  = (int) round($src_width * $ratio);
        $new_height = (int) round($src_height * $ratio);
        // สร้าง canvas ใหม่
        $canvas = imagecreatetruecolor($new_size, $new_size);
        // พื้นหลังสีขาว
        $white = imagecolorallocate($canvas, 255, 255, 255);
        imagefill($canvas, 0, 0, $white);
        //จัดกึ่งกลาง
        $dst_x = (int) (($new_size - $new_width) / 2);
        $dst_y = (int) (($new_size - $new_height) / 2);

        // Resize
        imagecopyresampled(
          $canvas,
          $src,
          $dst_x,
          $dst_y,
          0,
          0,
          $new_width,
          $new_height,
          $src_width,
          $src_height
        );

        // Save เป็น JPG
        if (!imagejpeg($canvas, $target_file, 90))
        {
          @unlink($source_file);
          echo 'Cannot save image';
          exit;
        }
      } // foreach ($use_size as $size)

      @unlink($source_file);

      $arr = array(
        'id' => $id_image,
        'style' => $code,
        'cover' => $this->product_image_model->has_cover($code) ? 0 : 1
      );

      $this->product_image_model->add($arr);
    }

    echo 'success';
  }

  public function set_cover_image()
  {
    $id = $this->input->post('id_image');
    $code = $this->input->post('style');
    $this->product_image_model->unset_cover($code);
    $this->product_image_model->set_cover($id);
    echo 'success';
  }

  public function remove_image()
  {
    $id = $this->input->post('id_image');
    $code = $this->input->post('style');

    if (! $this->product_image_model->delete_product_image($id))
    {
      echo 'fail';
      return;
    }

    if (! delete_product_image($id, $code))
    {
      echo 'fail';
      return;
    }

    echo 'success';
  }


  public function getImageSizeProperties($size)
  {
    $sc = array();
    switch ($size)
    {
      case "mini":
        $sc['prefix']  = "product_mini_";
        $sc['size']   = 60;
        break;
      case "default":
        $sc['prefix']   = "product_default_";
        $sc['size']   = 125;
        break;
      case "medium":
        $sc['prefix']   = "product_medium_";
        $sc['size']   = 250;
        break;
      case "large":
        $sc['prefix']   = "product_large_";
        $sc['size']   = 1500;
        break;
      default:
        $sc['prefix']   = "";
        $sc['size']   = 300;
        break;
    } //--- end switch
    return $sc;
  }
} //---- end class
