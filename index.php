<?php
/**
 * @package Posts2Excel
 * @version 1.0
 */
/*
Plugin Name: Posts2Excel
Plugin URI: http://www.mimasoftware.com/2014/04/posts2excel-wordpress-plugin.html
Description: Using this plugin you can store all your post into an excel spreadsheet for safe keeping.
Author: David Gallie
Version: 1.0
Author URI: http://mimasoftware.com
*/

/**
 * Setup our admin pages
 **/
 add_action("admin_menu", "p2e_create_admin_menu"); 
 
//1-parent, 2-page title,3-menu title,4-capability,5-menu slug(unique identifier),6-the function that handles and renders the page
function p2e_create_admin_menu(){
add_menu_page('Plugin settings', 'Posts2Excel', 'manage_options',
        'posts2excel_settings', 'p2e_posts2excel_settings_page');

}

/**
 *This function does all the grunt work for the main_admin_page.
 **/
function p2e_posts2excel_settings_page()
{
	/**
	 *Include phpexcel and render our little ad header regardless.
	 **/
	 include("PhpExcel/PHPExcel.php");
	 p2e_mimaheader();
	 /**
	  *Now we check if the form has been submitted
	  **/
	  if(isset($_REQUEST['submit']))
	  {
		  //forms been submitted so lets get to work
		 	$dlink = p2e_dumpposts(); //$dlink will either contain the url to the file or an error message if something failed
			$html = '<div id="downloadlink" class="updated settings-error">';
			$html .= '<div id="imgdl" class="imgdl"></div>';
			$html .= '<div id="dlink" class="dlink">'.$dlink.'</div>';
			$html .= '</div>';
			echo $html;
	  }
		  //form hasn't been submitted so show the user some options
		$form = '<form method="post" action="">
		<table class="form-table">
			<tr valign="top">
			<th scope="row"><label>Save Posts: </label></th>
				<td> 
					<input type="submit" class="button button-primary" name="submit" value="Go" />
				</td>
			</tr>
		</table>
			
			</form></div><!--./wrap-->';
		echo $form;
	  
}

function p2e_mimaheader()
{
	$html .='<div class="wrap">';
	$html .= '<h2>Posts2Excel</h2>';
	$html .= '<div id="p2e_mimaheader" class="updated settings-error">';
	$html .= 'Please visit <a href="http://www.mimasoftware.com/2014/04/posts2excel-wordpress-plugin.html">mimasoftware</a> if you need any help or support for this plugin, or if you would like to make a suggestion on how I can make it better.';
	$html .= '</div>';	
	
	echo $html;
}

function p2e_dumpposts()
{
	/*Create our spreadsheet and set its info and headers*/
$exobj = new PHPExcel();

$exobj->getProperties()->setCreator("Posts 2 Excel By Mimasoftware")
                             ->setLastModifiedBy("Mimasoftware")
                             ->setTitle("Posts 2 Excel Dump")
                             ->setSubject("Posts 2 Excel Dump")
                             ->setDescription("Lists all posts from the wordpress install")
                             ->setKeywords("Posts 2 Excel")
                             ->setCategory("");	
							 
/*add the headers to the sheet*/
$exobj->setActiveSheetIndex(0)
            ->setCellValue('A1', 'ID')
            ->setCellValue('B1', "post_author")
            ->setCellValue('C1', "post_date")
            ->setCellValue('D1', "post_date_gmt")
			->setCellValue('E1', "post_content")
			->setCellValue('F1', "post_title")
			->setCellValue('G1', "post_excerpt")
			->setCellValue('H1', "post_status")
			->setCellValue('I1', "comment_status")
			->setCellValue('J1', "ping_status")
			->setCellValue('K1', "post_password")
			->setCellValue('L1', "post_name")
			->setCellValue('M1', "to_ping")
			->setCellValue('N1', "pinged")
			->setCellValue('O1', "post_modified")
			->setCellValue('P1', "post_modified_gmt")
			->setCellValue('Q1', "post_content_filtered")
			->setCellValue('R1', "post_parent")
			->setCellValue('S1', "guid")
			->setCellValue('T1', "menu_order")
			->setCellValue('U1', "post_type")
			->setCellValue('V1', "post_mime_type")
			->setCellValue('W1', "comment_count")
			;
			
global $wpdb;
//$posts = $wpdb->get_row('SELECT * FROM wp_posts', ARRAY_A);
$posts = $wpdb->get_results( 'SELECT * FROM wp_posts',ARRAY_A );
$counter = 2;
foreach($posts as $post)
{

$exobj->setActiveSheetIndex(0)
            ->setCellValue('A'.$counter, $post['ID'])
            ->setCellValue('B'.$counter, $post["post_author"])
            ->setCellValue('C'.$counter, $post["post_date"])
            ->setCellValue('D'.$counter, $post["post_date_gmt"])
			->setCellValue('E'.$counter, $post["post_content"])
			->setCellValue('F'.$counter, $post["post_title"])
			->setCellValue('G'.$counter, $post["post_excerpt"])
			->setCellValue('H'.$counter, $post["post_status"])
			->setCellValue('I'.$counter, $post["comment_status"])
			->setCellValue('J'.$counter, $post["ping_status"])
			->setCellValue('K'.$counter, $post["post_password"])
			->setCellValue('L'.$counter, $post["post_name"])
			->setCellValue('M'.$counter, $post["to_ping"])
			->setCellValue('N'.$counter, $post["pinged"])
			->setCellValue('O'.$counter, $post["post_modified"])
			->setCellValue('P'.$counter, $post["post_modified_gmt"])
			->setCellValue('Q'.$counter, $post["post_content_filtered"])
			->setCellValue('R'.$counter, $post["post_parent"])
			->setCellValue('S'.$counter, $post["guid"])
			->setCellValue('T'.$counter, $post["menu_order"])
			->setCellValue('U'.$counter, $post["post_type"])
			->setCellValue('V'.$counter, $post["post_mime_type"])
			->setCellValue('W'.$counter, $post["comment_count"])
			;
			$counter += 1;
	}
	
/*set our sheet title and prompt the user to download the file*/
$exobj->getActiveSheet()->setTitle('Posts');
$exobj->setActiveSheetIndex(0);


// Redirect output to a client’s web browser (Excel5)
//header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
//header('Content-Disposition: attachment;filename="wp-posts.xlsx"');
//header('Cache-Control: max-age=0');
$sitename = get_bloginfo('name');
//$objWriter = PHPExcel_IOFactory::createWriter($exobj, 'Excel2007');
$objWriter = new PHPExcel_Writer_Excel2007($exobj);
$objWriter->save(plugin_dir_path( __FILE__ ).'/'.$sitename.'-posts.xlsx');  

//return a url to the newly created file
$dlink = plugin_dir_url(__FILE__).'/'.$sitename.'-posts.xlsx';

$output = "";
//check the file has been written and saved properly before offering a link to download it
if(file_exists(plugin_dir_path(__FILE__).'/'.$sitename.'-posts.xlsx'))
{
	$output .= '<p>Your posts have been successfully saved! Click the link below to download:</p>';
	$output .= '<a href="'.$dlink.'">'.$sitename.'-posts.xlsx'.'</a>';
}else{
	$output .= '<p>There seems to have been a problem saving the Excel file. Please click the link at the top of this page for help in resolving this problem.</p>';
}

return $output;
}




   
?>