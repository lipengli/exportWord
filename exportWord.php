<?php
class word
{ 
	private $xmlns = '<html xmlns:o="urn:schemas-microsoft-com:office:office"
		xmlns:w="urn:schemas-microsoft-com:office:word"
		xmlns="http://www.w3.org/TR/REC-html40">';
		
	/**
	*	���캯��
	*/
	
	public function __construct(){

	}
	/**
	*	�޵�����ʾstart
	*/
	public function start()
	{
		ob_start();
		echo $this->xmlns;
	}
	
	/**
	*	�޵�����ʾ����
	*/
	public function save($path)
	{
		echo "</html>";
		$data = ob_get_contents();
		ob_end_clean();
		 
		$this->wirteFile($path,$data);
	}
	

	 
	/**
	*	�޵�����ʾ����
	*/
	function wirteFile($fn,$data)
	{
		$fp=fopen($fn,"wb");
		fwrite($fp,$data);
		fclose($fp);
	}
	
	/**
	*	�е�����ʾstart
	*/
	public function flipStart($fileName = 'lpFuture.doc'){
		ob_start();
		header("Content-type: application/octet-stream;charset=gbk");
		header("Accept-Ranges: bytes");
		header("Content-Disposition: attachment; filename=".$fileName);
		echo $this->xmlns;
	}
	
	/**
	*	�е�����ʾ����
	*/
	public function flipSave($path)
	{
		echo "</html>";
		$data = ob_get_contents();
		ob_end_clean();
		 
		$this->flipWirteFile($data);
	}
	
	/**
	*	�е�����ʾ����
	*/
	function flipWirteFile($data)
	{
		echo $data;
	}
}


//���Դ��룺
$html = ' 
<h3>��һ��Ԫ<h3>
<div>1-1 .3<br><span>&nbsp;&nbsp;�����ԡ�</span>Ӧ�ɶ�δ�ɵĳ��ʶ�ȴﵽ<span>��10000��</span>Ԫ,������<span>��12��</span>�¡�<spanclass="extras_span_hidden_button"data-unit_extras="_unit_1_1_3_extras">+</span><Br><span>������1��</span>Ӧ�ɶ�δ�ɵĳ��ʶ�ȴﵽ<span>��10040002003��</span>Ԫ,������<span>��11��</span>�¡�<spanclass="extras_span_hidden_del"data-unit_extras="_unit_1_1_3_extras">ɾ��</span></div><div>1-1 .6<br><span>�����ԡ�</span>��<span>��2015��</span>��<span>��12��</span>��<span>��22��</span>����<span>��������</span>ת��<span>��100��</span>%�Ĺ�Ȩ��ת��������δ������ϣ���Ҫԭ����<span>����Ǯ�����ԡ�</span><Br></div><div>1-1 .7<br>��������ľ�������Ϊ<span>��̫��Ǯ������ν��̬�ȡ�</span><Br></div><div>1-1 .8<br>�����������¹�<span>�����ơ�</span><Br></div><div>1-1 .10<br>��������δ��ɻ���ҵЭ��ĵǼ�������<br></div><div>1-1 .11<br><span>��2015��</span>����<span>��̫��Ǯ��</span>������첻�ϸ�<br><span>��2016��</span>����<span>��Ǯ������ȥ��</span>���±����̵Ǽ����ܻ�����ʾ��Ӫ�쳣��<spanclass="extras_span_hidden_button"data-unit_extras="_unit_1_1_11_A_extras">+</span><Br></div><div>1-1 .12<br>��<span>����Ǯ��</span>����������<span>��2015��</span>��<span>��2��</span>��<span>��12��</span>���ܵ�<span>��������</span>�����<span>��100000000000����Ԫ��</span>������������ʽΪ<span>�����</span>��<spanclass="extras_span_hidden_button"data-unit_extras="_unit_1_1_12_extras">+</span><Br></div>
'; 
 
//�������� 
for($i=1;$i<=3;$i++){ 
    $word = new word(); 
    $word->flipStart(); 
    //$html = "aaa".$i; 
    $wordname = time().$i.".doc"; 
    echo $html; 
    $word->flipSave($wordname); 
    ob_flush();//ÿ��ִ��ǰˢ�»��� 
    flush(); 
}