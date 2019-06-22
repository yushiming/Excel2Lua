package application;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.ResourceBundle;

import com.sun.javafx.robot.impl.FXRobotHelper;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Button;
import javafx.scene.control.ListView;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

public class MyController implements Initializable {

   @FXML
   private Button _btnOpen;             //打开Excel所在目录
   @FXML
   private Button _btnTranf;            //开始转换
   @FXML
   private Button _btnOutDir;           //打开转换后的目录
   @FXML
   private ListView _listView;          //显示Excel表文件列表
   @FXML
   private TextArea _textArea;          //转换过程控制台输出
   
   
   private String _outDir;              //输出文件夹
   
   
   /**
    * listview数据
    */
   private ObservableList<String> _excelList = FXCollections.observableArrayList();

   
   @Override
   public void initialize(URL location, ResourceBundle resources) {

       // TODO (don't really need to do anything here).
	   
	   _btnTranf.setDisable(true);
	   _btnOutDir.setDisable(true);
	   
   }
   
   public void openExcelDirDialog(ActionEvent event) {
       DirectoryChooser directoryChooser = new DirectoryChooser();
       directoryChooser.setTitle("选择导出Excel文件夹");
       ObservableList<Stage> stage = FXRobotHelper.getStages();

       File dirFile = directoryChooser.showDialog(stage.get(0));
       if(dirFile != null ) {
    	   String dir = dirFile.getAbsolutePath();
    	   String outDirName = dirFile.getName() + "out";
    	   String workDir = dirFile.getParent();
    	   _outDir = workDir + "\\" + outDirName;
    	   System.out.println("_outDirName : " + outDirName);
    	   System.out.println("_workDir : " + workDir);
           File[] subfiles = dirFile.listFiles();
    	   _listView.setItems(_excelList);
           for (File file : subfiles) {
        	   _excelList.add(file.getAbsolutePath());
        	   System.out.println("file : " + file.getAbsolutePath());
           }
    	 
    	   _btnTranf.setDisable(false);
       }
       else
       {
    	   _btnTranf.setDisable(true);
       }     
   }
               
   public void tranfToLua(ActionEvent event) {
	   
		// 指定路径如果没有则创建并添加
		File dir = new File(_outDir);
		//判断是否存在
		if (!dir.exists()) {
		//创建目录文件
			dir.mkdirs();
		}
		
		File[] subfiles = dir.listFiles();//取得当前目录下所有文件和文件夹
		for (File file : subfiles) {
			if(file.isFile()){//判断是否是文件
				file.delete();
			}
		}
		
		String outStr = "输出目录: " + _outDir + "\n";
		//_textArea.setText(outStr);
	   
	    for(int i = 0; i < _excelList.size(); i++){
		   	String infile = _excelList.get(i);
		   
			String path = infile.substring(0, infile.lastIndexOf("\\"));
			String fileName = infile.substring(infile.lastIndexOf("\\")+1);  
			String fileNameWithNoEx = fileName.substring(0, fileName.indexOf('.'));
			String outLuaPath = _outDir + "\\" + fileNameWithNoEx + ".lua";
			
			System.out.println("infile : " + infile);
			
			outStr += outLuaPath + "\n";
			//_textArea.setText(outStr);
			
			Helper.excel2Lua(infile , outLuaPath, fileNameWithNoEx);
	    }
	   
	    _btnOutDir.setDisable(false);
	    outStr = outStr + "导出成功！";
	    _textArea.setText(outStr);
   }
   
   // When user click on myButton
   // this method will be called.
   public void showOutDir(ActionEvent event) {
	   try {
		Runtime.getRuntime().exec("explorer " + _outDir);
	} catch (IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
	   //Desktop.getDesktop().open(new File("文件路径"));
   }
   
   
   
   
}




