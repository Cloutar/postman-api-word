<?php
/**
 * postman api 例子类
 */
namespace App\Libraries;

use App\Http\Controllers\Controller;

class PostmanApiEG {

    public $http_method;
    public $eg_url;
    public $request_title;
    public $request_content;
    public $response_title;
    public $response_content;


    public function __construct($http_method,
                                $eg_url ,
                                $request_title,
                                $request_content,
                                $response_title,
                                $response_content) {
        $this->http_method = $http_method;
        $this->eg_url = $eg_url;
        $this->request_title = $request_title;
        $this->request_content = $request_content;
        $this->response_title = $response_title;
        $this->response_content = $response_content;
    }


    /**
     * 生成例子表格
     * @param $php_word [PHPWord的实例]
     * @param $section [页面]
     */
    public function makePhpWordForm($php_word, $section) {
        $controller = new Controller();

        //表格样式
        $table_title_style = ['bgColor' => '000000'];
        $php_word->addTableStyle('tableStyle', ['borderColor' => '000000', 'borderSize' => 6, 'cellMargin' => 50]);

        $table = $section->addTable('tableStyle');
        $table->addRow();
        $table->addCell(null, $table_title_style)->addText( $controller->filterOutXMLIllicit($this->http_method) );
        $table->addRow();
        $table->addCell()->addText( $controller->filterOutXMLIllicit($this->eg_url) );

        $table->addRow();
        $table->addCell(null, $table_title_style)->addText( $controller->filterOutXMLIllicit($this->request_title) );
        $table->addRow();
        $table->addCell()->addText(  $controller->jsonBreak2wordBreak( $controller->filterOutXMLIllicit($this->request_content) ) );

        $table->addRow();
        $table->addCell(null, $table_title_style)->addText( $controller->filterOutXMLIllicit($this->response_title) );
        $table->addRow();
        $table->addCell()->addText( $controller->jsonBreak2wordBreak( $controller->filterOutXMLIllicit($this->response_content) ) );

        $table->setWidth(4500);
    }
}
