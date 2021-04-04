<?php

namespace App\Http\Controllers;

use App\Libraries\PostmanApiEG;   //postman api例子文档 类
use Illuminate\Foundation\Auth\Access\AuthorizesRequests;
use Illuminate\Foundation\Bus\DispatchesJobs;
use Illuminate\Foundation\Validation\ValidatesRequests;
use Illuminate\Http\Request;
use Illuminate\Routing\Controller as BaseController;
use Illuminate\Support\Facades\Validator;
use Illuminate\Validation\ValidationException;
use PhpOffice\PhpWord\Element\Section;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\PhpWord;

class Controller extends BaseController
{
    use AuthorizesRequests, DispatchesJobs, ValidatesRequests;

    /**
     * ['生成 Postman API 文档']
     * @param Request $request
     * @return \Illuminate\Http\JsonResponse
     * @throws ValidationException
     * @throws \Illuminate\Contracts\Filesystem\FileNotFoundException
     * @throws \PhpOffice\PhpWord\Exception\Exception
     */
    public function postmanApiWord(Request $request) {
        $params = $request->all();
        $rules = [
            'postman_api_doc' => 'bail|required|file',
        ];
        $messages = [
            'postman_api_doc.required' => '必须传递postman_api_doc参数',
            'postman_api_doc.file' => 'postman_api_doc参数必须为成功上传的文件'
        ];
        $validator = Validator::make($params, $rules, $messages);

        if ($validator->fails()) {
            $messages = $validator->errors()->first();
            throw new ValidationException($validator, response()->json([
                'code' => 1,
                'message' => $messages,
            ]), 422);
        }

        //处理上次文件中的内容
        $postman_api = json_decode( $request->file('postman_api_doc')->get(), true );

        //实例化
        $php_word = new PhpWord();

        //添加页面
        $section = $php_word->addSection();

        //标题
        $php_word->addTitleStyle(1, ['bold' => true, 'size' => 42], ['alignment' => 'center']);
        $section->addTitle($postman_api['info']['name'], 1);
        $section->addTextBreak();

        //一级标题
        $php_word->addTitleStyle(2, ['bold' => true, 'size' => 38]);

        //文档内容递归输出，最后两个参数由上面得出, 最后一个参数是默认文本大小得出
        $this->loop($section, $postman_api['item'], $php_word, 2, 38, 12);

        //文档结束
        $section->addTextBreak();
        $section->addText('（以下空白）', ['fontName' => '微软雅黑', 'fontSize' => 12]);

        //生成Word2010文档
        $writer = IOFactory::createWriter($php_word);
        $writer->save(storage_path('app/postmanWord/postman-api.docx'));

        return response()->json([
            'code' => 0,
            'message' => 'postman api文档转换为word文档',
        ]);
    }

    /**
     * [递归处理postman的接口数据]
     * @param Section $section [phpword的页面]
     * @param $item [postman文档的中的item部分]
     * @param $php_word [phpword实例]
     * @param int $depth [标题的级数, 初始级数为2(一级标题)]
     * @param int $size [字体大小, 初始为38(一级标题), 每个阶级递减4]
     * @param int $min_size [最小字体大小, 默认为12]
     * @param int $indent [文本缩进, 默认值为0(一级标题)]
     * @author Cloutar
     * @date 2021-03-31
     */
    public function loop(Section $section, $item, $php_word, $depth = 2, $size = 38, $min_size = 12, $indent = 0) {
        foreach ($item as $key => $value) {
            //文件夹
            if (isset($value['item'])) {
                $section->addTitle($this->filterOutXMLIllicit($value['name']), $depth);
                //下一个标题
                $size = ($size - 4) <= $min_size ? $min_size : ($size - 4);
                $php_word->addTitleStyle($depth + 1, ['bold' => true, 'size' => $size], ['indent' => $indent + 1]);
                //回调
                $this->loop($section, $value['item'], $php_word, $depth + 1, $size - 4, $min_size, $indent + 1);
            }
            //接口
            else {
                $section->addTitle('-'. $this->filterOutXMLIllicit($value['name']), $depth);
                $section->addText( '接口地址：' . $this->filterOutXMLIllicit($value['request']['url']['raw']), ['fontName' => '微软雅黑', 'fontSize' => $size - 4], ['indent' => $indent + 1]);
                $section->addText('接口方法：' . $this->filterOutXMLIllicit($value['request']['method']), ['fontName' => '微软雅黑', 'fontSize' => $size - 4], ['indent' => $indent + 1]);

                $section->addText('调用例子：', ['fontName' => '微软雅黑', 'fontSize' => $size - 4], ['indent' => $indent + 1]);
                //get query 情况
                if (isset($value['response'][0]['originalRequest']['url']['query'])) {
                    $query = '';
                    foreach ($value['response'][0]['originalRequest']['url']['query'] as $key => $item) {
                        if ($key <> 0) {
                            $query .=  "\n";
                        }
                        $query .= $this->filterOutXMLIllicit($item['key']) . ' : ' . $this->filterOutXMLIllicit($item['value']);
                    }
                    //生成例子表格
                    (new PostmanApiEG(
                        $value['response'][0]['originalRequest']['method'] . ' : ',
                        $value['response'][0]['originalRequest']['url']['raw'],
                        'Query Params: ',
                        $query,
                        'Response(' . $value['response'][0]['code'] . '): ',
                        $value['response'][0]['body']))->makePhpWordForm($php_word, $section);
                }
                //POST RAW JSON 情况
                else if (isset($value['response'][0]['originalRequest']['body']['options']['raw']['language']) && $value['response'][0]['originalRequest']['body']['options']['raw']['language'] == 'json') {
                    (new PostmanApiEG(
                        $value['response'][0]['originalRequest']['method'] . ' : ',
                        $value['response'][0]['originalRequest']['url']['raw'],
                        'Request(RAW-JSON): ',
                        $value['response'][0]['originalRequest']['body']['raw'],
                        'Response(' . $value['response'][0]['code'] . '): ',
                        $value['response'][0]['body']))->makePhpWordForm($php_word, $section);
                }
                else {
                    $section->addText('(暂无)', ['fontName' => '微软雅黑', 'fontSize' => $size - 4], ['indent' => $indent + 2]);
                }
            }
            //空行
            $section->addTextBreak();
        }
    }


    /**
     * [正则转换, 防止word文档发生错误]
     * @param $string [字符串，并兼容null]
     * @return string
     * @author Cloutar
     * @date forget
     */
    public function filterOutXMLIllicit($string) :string {
        $patterns = [
            '/&/',
            '/ /',
            '/</',
            '/>/',
            '/"/',
            '/\'/',
        ];
        $replacements = [
            '&#38;',
            '&#160;',
            '&#60;',
            '&#62;',
            '&#34;',
            '&#39;',
        ];

        return preg_replace($patterns, $replacements, $string);
    }

    /**
     * [让\n换行符号转化为word文档的换行符号]
     * @param $string [字符串，并兼容null]
     * @return string
     * @author Cloutar
     * @date 2021-03-31
     */
    public function jsonBreak2wordBreak($string) :string {

        return str_replace("\n", '<w:br />', $string);
    }

}
