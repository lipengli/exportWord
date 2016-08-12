<?php
class WordoneHelper
{ 
	private $xmlns = '<html xmlns:o="urn:schemas-microsoft-com:office:office"
		xmlns:w="urn:schemas-microsoft-com:office:word"
		xmlns="http://www.w3.org/TR/REC-html40"><head><xml><w:WordDocument><w:View>Print</w:View></xml></head>';
	//<head><xml><w:WordDocument><w:View>Print</w:View></xml></head> 普通格式打开word  如果不加，则是web格式打开
		
	/**
	*	无弹层提示start
	*/
	private function start()
	{
		ob_start();
		echo $this->xmlns;
	}
	
	/**
	*	无弹层提示保存
	*/
	private function save($path)
	{
		echo "</html>";
		$data = ob_get_contents();
		ob_end_clean();
		 
		$this->wirteFile($path,$data);
	}
	

	 
	/**
	*	无弹层提示保存
	*/
	private function wirteFile($fn,$data)
	{
		$fp=fopen($fn,"wb");
		fwrite($fp,$data);
		fclose($fp);
	}
	
	/**
	*	有弹层提示start
	*/
	private function flipStart($fileName = 'lpFuture.doc'){
		ob_start();
		header("Content-type: application/octet-stream;charset=gbk");
		header("Accept-Ranges: bytes");
		header("Content-Disposition: attachment; filename=".$fileName);
		echo $this->xmlns;
	}
	
	/**
	*	有弹层提示保存
	*/
	private function flipSave()
	{
		echo "</html>";
		$data = ob_get_contents();
		ob_end_clean();
		 
		$this->flipWirteFile($data);
	}
	
	/**
	*	有弹层提示保存
	*/
	private function flipWirteFile($data)
	{
		echo $data;
	}
	
	/**
	*	外部使用方法：导出word
	*	@params string $html 导出内容
	*	@params string $path 导出的地址
	*/
	public function write($html,$path){
		$wordname = time().rand(1,10).".doc"; 
		$this->flipStart($wordname); 
		//$html = "aaa".$i; 
		echo $html; 
		$this->flipSave(); 
		ob_flush();//每次执行前刷新缓存 
		flush(); 
	}
}


//测试代码：
$html = ' 
<h3>第一单元<h3>
<div>1-1 .3<br><span>&nbsp;&nbsp;【测试】</span>应缴而未缴的出资额度达到<span>【10000】</span>元,已逾期<span>【12】</span>月。<spanclass="extras_span_hidden_button"data-unit_extras="_unit_1_1_3_extras">+</span><Br><span>【测试1】</span>应缴而未缴的出资额度达到<span>【10040002003】</span>元,已逾期<span>【11】</span>月。<spanclass="extras_span_hidden_del"data-unit_extras="_unit_1_1_3_extras">删除</span></div><div>1-1 .6<br><span>【测试】</span>于<span>【2015】</span>年<span>【12】</span>月<span>【22】</span>日向<span>【李鹏】</span>转让<span>【100】</span>%的股权。转让手续尚未履行完毕，主要原因是<span>【有钱、任性】</span><Br></div><div>1-1 .7<br>存在争议的具体情形为<span>【太有钱，无所谓的态度】</span><Br></div><div>1-1 .8<br>合作方还从事过<span>【慈善】</span><Br></div><div>1-1 .10<br>合作方尚未完成基金业协会的登记手续。<br></div><div>1-1 .11<br><span>【2015】</span>年因<span>【太有钱】</span>导致年检不合格。<br><span>【2016】</span>年因<span>【钱花不出去】</span>导致被工商登记主管机关提示经营异常。<spanclass="extras_span_hidden_button"data-unit_extras="_unit_1_1_11_A_extras">+</span><Br></div><div>1-1 .12<br>因<span>【有钱】</span>，合作方于<span>【2015】</span>年<span>【2】</span>月<span>【12】</span>日受到<span>【李鹏】</span>给予的<span>【100000000000亿万元】</span>处罚，处罚方式为<span>【罚款】</span>。<spanclass="extras_span_hidden_button"data-unit_extras="_unit_1_1_12_extras">+</span><Br></div>
'; 
 
//批量生成 
for($i=1;$i<=3;$i++){ 
    $word = new word(); 
    $word->write($html);
}
