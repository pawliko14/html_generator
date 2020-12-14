import java.awt.Desktop;
import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import java.awt.Font;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;

import java.io.File;
import java.io.IOException;


public class plc_alarm_gui {

	private JFrame frame;
	private JTextField txtPodajPlikXls;
	private JTextField txtProgramDoTworzenia;
	private JTextField txtPodajNazweFolderu;
	private JTextField textField;
	private JButton btnPlik;
	private String filename;
	public static  String nazwa;
	private JButton btnOtworzFolder;
	private  File file;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					plc_alarm_gui window = new plc_alarm_gui();
							
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public plc_alarm_gui() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.getContentPane().setFont(new Font("Tahoma", Font.BOLD, 11));
		frame.setBounds(100, 100, 450, 300);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		
		txtPodajPlikXls = new JTextField();
		txtPodajPlikXls.setText("Podaj plik xls z alarmami");
		txtPodajPlikXls.setBounds(63, 75, 173, 20);
		txtPodajPlikXls.setEnabled(false);
		frame.getContentPane().add(txtPodajPlikXls);
		txtPodajPlikXls.setColumns(10);
		
		txtProgramDoTworzenia = new JTextField();
		txtProgramDoTworzenia.setFont(new Font("Tahoma", Font.BOLD, 13));
		txtProgramDoTworzenia.setText("program do tworzenia alarmow w odrebnych plikach html");
		txtProgramDoTworzenia.setEditable(false);
		txtProgramDoTworzenia.setBounds(24, 11, 389, 37);
		txtProgramDoTworzenia.setEnabled(false);
		frame.getContentPane().add(txtProgramDoTworzenia);
		txtProgramDoTworzenia.setColumns(10);
		
		txtPodajNazweFolderu = new JTextField();
		txtPodajNazweFolderu.setText("podaj nazwe folderu do zapisu");
		txtPodajNazweFolderu.setColumns(10);
		txtPodajNazweFolderu.setBounds(63, 122, 173, 20);
		txtPodajNazweFolderu.setEnabled(false);
		frame.getContentPane().add(txtPodajNazweFolderu);
		
	
		btnOtworzFolder = new JButton("Otworz folder");
		btnOtworzFolder.setVisible(false);
		
		btnOtworzFolder.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
			       Desktop desktop = Desktop.getDesktop();
			        File dirToOpen = null;
			        try {
			            dirToOpen = new File( nazwa);
			            desktop.open(dirToOpen);
			        } catch (IllegalArgumentException iae) {
			            System.out.println("File Not Found: "+ nazwa);
			        } catch (IOException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
			    
			}
		});
		btnOtworzFolder.setBounds(168, 206, 136, 23);
		frame.getContentPane().add(btnOtworzFolder);
		
		
//		textField = new JTextField();
//		textField.setBounds(271, 122, 119, 20);
//		frame.getContentPane().add(textField);
//		textField.setColumns(10);
//		
//		nazwa = textField.getText();
		   
		JButton btnFIleDestination = new JButton("destination");
			btnFIleDestination.setBounds(271, 122, 119, 20);
			frame.getContentPane().add((btnFIleDestination));
			
			
			
			
			
			btnFIleDestination.addActionListener(new ActionListener() {
				public void actionPerformed(ActionEvent arg0) 
				{
				     JFileChooser fileChooser = new JFileChooser();
			            fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
			            int option = fileChooser.showOpenDialog(frame);
			            
			            if(option == JFileChooser.APPROVE_OPTION)
			            {
			               file = fileChooser.getSelectedFile();
			               
			               nazwa = file.getAbsoluteFile().toString();
			               txtPodajNazweFolderu.setText(nazwa);
			            }
			            
					
				}
			});
			
			
			
			
			
			
			
			
			
			////
			
			
		btnPlik = new JButton("plik");
		
		
		
		
		
		
		 ////////////
		
		JButton btnNewButton = new JButton("utworz pliki HTML");
		
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
			
				try {
				//	nazwa = textField.getText();
					SourceCode.RUN(filename,nazwa);
					btnOtworzFolder.setVisible(true);

				} catch (EncryptedDocumentException | InvalidFormatException | IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
			        JOptionPane.showMessageDialog(null, e1, "prawodpodobnie plik excela nei jest zamkniety, zamknij ", JOptionPane.INFORMATION_MESSAGE);
				}
				
			}
		});
		btnNewButton.setBounds(61, 172, 329, 23);
		frame.getContentPane().add(btnNewButton);
		
		btnPlik.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser chooser = new JFileChooser();
				FileNameExtensionFilter filter = new FileNameExtensionFilter("xls","xls");
				chooser.setFileFilter(filter);
				chooser.showOpenDialog(null);
				
				File f = chooser.getSelectedFile();
				 filename = f.getAbsolutePath();
				
				//path.setText(filename);
				
				 txtPodajPlikXls.setText(filename);
			}
		});
		btnPlik.setBounds(271, 74, 113, 23);
		frame.getContentPane().add(btnPlik);
		
		
	
	}
}
