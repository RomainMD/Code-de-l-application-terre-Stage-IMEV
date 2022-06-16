using System;
using System.Threading;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Data.SqlClient;//sqldata
using Excel = Microsoft.Office.Interop.Excel;



namespace App_1
{
	

	public partial class Form1 : Form
	{

		DataTable DatatableFLUNTR = new DataTable();//
		DataTable DatatableFLUNTR1 = new DataTable();
		DataTable DatatableFLUNTR2 = new DataTable();
		DataTable DatatableFLUNTR3 = new DataTable();


		//alphabet de decodage
		ASCIIEncoding ascii;

		// buffer de reception
		int sum;
		string newLine = Environment.NewLine;
		string dataascii1;
		string dataascii2;
		string dataascii3;
		int somme1;
		int somme2;
		int somme3;
		int j=0;
		int cptligne = 0;//ou 1 ?
		int nbrligne1;
		int nbrligne2;
		int nbrligne3;
		string checksumfiltrer;//checksum recuperer/recu
		string valCS1;
		string valCS2;
		string valCS3;
		int k;
		string nbrligne;//recu !
		string valnbrL1;
		string valnbrL2;
		string valnbrL3;
		string sommevrai;
		int l;
		string nbrlignefiltrer;
		const int MAXDATA = 1048576;//1024
		byte[] MyReceiveBuffer;
		byte[] myReceiveData;
		char[] tabdata;
		char[] tab;
		bool FILTRE = true; //bool TEST = true;
		bool TESTk = true;
		bool TESTl = true;
		byte sumdata;
		String compteurc;
		String cptaqui;
		String dataascii;
		String dataasciivrai;
		String datahorloge;
		bool verificationchecksomme = false;
		bool verificationnbrligne = false;
		int checksommecalculervrai;
		int nbrlignecalculervrai;
		string checksommerecuvrai;
		string nbrlignerecuvrai;
		bool modeauto = false;//mode auto desactiver par defaut
		byte oldcara;
		byte oldoldcara;
		byte _3oldcara;
		bool testHorloge;
		string unelignedonnees;
		char[] tablignedonnees;
		int cpts = 0;
		bool filtredonnee;
		String strDate;
		String strHeure;
		String strTemperature;
		String strLambda1;
		String strCHL;
		String strLambda2;
		String strNTU;
		String strThermistance;
		String @datatxtpathfile; //l'emplacement de fichier et enregistrer
		String heuresbouee;
		String minutesbouee;
		String secondesbouee;
		bool ReceptionAuto;
		bool testexcel;
		String[] Heureinterface;
		int r;
		int cpt;//compte les secondes
		DataRow workRow;
		int numerotramevrai;
		int ligneexcel;
		int joursansreception;

		private static System.Threading.Timer time1;


		Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();//a ne pas suppr^^

		public Form1() //pourquoi pas partial class form1:form ?//public Form1
		{
			//Application.Restart();

			InitializeComponent();
			CheckForIllegalCrossThreadCalls = false;
			serialPort1.Open();
			textBox8.Text = (serialPort1.BaudRate).ToString();
			MyReceiveBuffer = new byte[MAXDATA];//voir si pb possible
			// Construction de l'alphabet ASCII
			ascii = new ASCIIEncoding();
			serialPort1.Parity = Parity.None;
			serialPort1.DataBits = 8;
			serialPort1.StopBits = StopBits.One;
			time1 = new System.Threading.Timer(Timenow, null, 0, 1000);//toutes les 1 sec appelle  datetime

			@datatxtpathfile = Properties.Settings.Default.EmplacementFichier;
			label19.Text = @datatxtpathfile;
			label19.ForeColor = Color.Black;

			 ligneexcel = Properties.Settings.Default.nombredeligneexcel;
			/*	
				bg.Tick += (s, e) => { textBox10.Text = DateTime.Now.ToString(); };
				bg.Interval = 500;
				bg.Start();
			*/

			//////
			//   Properties.Settings.Default["SomeProperty"] = "Some Value";
			//   Properties.Settings.Default.Save(); //Saves settings in application configuration file
			/////
			///

		}

		private void Form1_Load(object sender, EventArgs e)
		{
			//?
			initDatatableFLUNTR1();
			initDatatableFLUNTR2();
			initDatatableFLUNTR3();
			serialPort1.Encoding = ASCIIEncoding.ASCII;
			// Enumeration des ports COM presents
			foreach (string sPortName in SerialPort.GetPortNames())
			{
				comboBox1.Items.Add(sPortName);
			}
			
			
		}
		public void Timenow(Object state)
        {
			//cpt++;
			textBox10.Clear();
			textBox10.Text = DateTime.Now.ToString();
			/*textBox10.AppendText(newLine + DateTime.Now.Day.ToString());//5 au lieu de 05
			textBox10.AppendText(newLine + DateTime.Now.Month.ToString());
			textBox10.AppendText(newLine + DateTime.Now.Year.ToString());
			textBox10.AppendText(newLine + DateTime.Now.Hour.ToString());
			textBox10.AppendText(newLine + DateTime.Now.Minute.ToString());
			textBox10.AppendText(newLine + DateTime.Now.Second.ToString());*/
			
			if (DateTime.Now.Hour == 16 && DateTime.Now.Minute == 54 && DateTime.Now.Second == 30)
			{
				if (modeauto == true)
				{
					initDatatableFLUNTR1();
					initDatatableFLUNTR2();
					initDatatableFLUNTR3();
					dataascii1 = "";
					dataascii2 = "";
					dataascii3 = "";
					textBox1.Clear();
					textBox2.Clear();
					textBox3.Clear();
					textBox6.Clear();
					textBox7.Clear();
					textBox9.Clear();
					j = 0;
					k = 0;
					l = 0;
					cptligne = 0;
					sum = 0;
					serialPort1.Write("t");
					testHorloge = true;
				}
			}
			if (DateTime.Now.Hour ==13 && DateTime.Now.Minute==35 && DateTime.Now.Second==00) {//
				if (modeauto==true) {
					if (joursansreception >= 5)
					{
						serialPort1.Write("v");
					}
					modeAUTO();
				}
			}
			//if timer2 () //if (cpt == 300)//5 min apres demande de reception donnees
			//{

		//	else if (DateTime.Now.Hour == 13 && DateTime.Now.Minute == 40 && DateTime.Now.Second == 00)//10 min plus tard
			//{
				/*if (modeauto == true)
				{
					//if (textBox9.Text =="")//s'il n'y a pas de texte
                    //{
						//textBox1.AppendText(newLine+DateTime.Now.Day + DateTime.Now.Hour + DateTime.Now.Minute + ": Rien n'a été recu");
						//
                    //}
					String pathfile = @datatxtpathfile;
					//String pathfile = @"C:\Users\grisoni\Desktop\";// @"C:\Users\ROMAIN.000\Desktop\"
					String filename = "data_FLUNTR.txt";
					System.IO.File.AppendAllText(pathfile + filename, textBox9.Text);//textbox2>>>9//on sauvegarde les données filtrer //a remplacer par la text box des donnes filtrer verifier quand le checksum sera OK
					textBox9.Clear();
				}*/
			//}

		}
		private void serialPort1_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
		{
			/*string serialread;
			string serialread2;

			serialread = serialPort1.ReadExisting();
			textBox1.AppendText(serialread);*/

			String decodedAscii;
			int iNbByteReceived;
			// On verifie le nombre d'octets reçus
			iNbByteReceived = serialPort1.BytesToRead;
			// Debug
			//textBox1.AppendText("Réception de " + iNbByteReceived.ToString() + " octets");//si on veut le nbr d'octet envoyé
			//textBox1.AppendText("debug");
			if (iNbByteReceived != 0) //utile ?
			{
				// On sort du buffer de reception du driver serie le nombre d'octets reçus
				// On place ces octets dans notre buffer tampon
				
				serialPort1.Read(MyReceiveBuffer, 0, iNbByteReceived);
				
				/*serialread = serialPort1.ReadExisting();
				textBox1.AppendText(serialread);
				serialread2 = serialPort1.ReadLine();
				textBox1.AppendText(serialread2);*/

				// On interprete le tableau d'octet comme une chaine de caractere ASCII
				decodedAscii = ascii.GetString(MyReceiveBuffer, 0, iNbByteReceived);
				// affichage de la chaine ascii dans textBox1
				textBox1.AppendText(decodedAscii);
				textBox1.SelectionStart = textBox1.Text.Length;
				textBox1.ScrollToCaret();
				textBox1.Refresh();
				//}//end if
				//tab = decodedAscii.ToCharArray();//
				// ou methode avec myReceivebuffer
				// textBox3.Text = textBox3.Text + iNbByteReceived;//pour debuguer 
				for (int i = 0; i < iNbByteReceived; i++){//ex rec "1<121>2" inb=7

					int testvalascii = MyReceiveBuffer[i];
					//textBox15.AppendText(" " + testvalascii + " ");//cetait pour aff les vaeurs en decimal des cara (ascii)
														  //textBox2.AppendText(newLine + MyReceiveBuffer[i]);//debug
					if ((MyReceiveBuffer[i] == '>') && (oldcara == '>') && (oldoldcara == '>')) //&& oldcara =='>'
					{
						FILTRE = true;
						j++;
						if (j == 1)
						{
							somme1 = sum; //on enregistre la valeur du checksomme calculer pour cette trame dans une variable
							textBox3.AppendText(newLine + "S1 = " + somme1.ToString());
							sum = 0;
							nbrligne1 = cptligne;
							cptligne = 0;
						}
						if (j == 2)
						{
							somme2 = sum;
							textBox3.AppendText(newLine + "S2 = " + somme2.ToString());
							sum = 0;
							nbrligne2 = cptligne;
							cptligne = 0;
						}
						if (j == 3)
						{
							somme3 = sum;
							textBox3.AppendText(newLine + "S3 = " + somme3.ToString());
							j = 0;
							sum = 0;
							nbrligne3 = cptligne;
							cptligne = 0;
						}
						//textBox2.Text = textBox2.Text + "   >   ";//pour afficher
					}
					if ((FILTRE == false) && ((oldcara == '>') && (MyReceiveBuffer[i] != '>')))
					{ sum += 62; }
					if ((FILTRE == false) && ((oldcara == '>') && (oldoldcara == '>')) && (MyReceiveBuffer[i] != '>'))
					{ sum += 62; }

					if ((FILTRE == false )&&(MyReceiveBuffer[i]!='>'))// quand on est apres <<< avant >>> sans compter les ">>"
					{
						/*
						int checksomme = CheckSum(MyReceiveBuffer[i]);////////////checksomme
						sum += checksomme;
						*/
						//byte = 0x00
						//for (int i = 0; i < byteData.Length; i++) //s'il y a plusieurs cara
						//chkSumByte ^= byteData[i];

						//if ((MyReceiveBuffer[i] != 13) || (MyReceiveBuffer[i] != 10)) //ne fait rien du tout lol; CR+LF //L'interface bouée ne prennant pas en compte les retour a la ligne dans son checksomme
						//{
							sum += MyReceiveBuffer[i];
						//}

						if ( (MyReceiveBuffer[i]==13) ||(MyReceiveBuffer[i] =='	') || (MyReceiveBuffer[i] == '0') || (MyReceiveBuffer[i] == '1') || (MyReceiveBuffer[i] == '2') || (MyReceiveBuffer[i] == '3') || (MyReceiveBuffer[i] == '4') || (MyReceiveBuffer[i] == '5') || (MyReceiveBuffer[i] == '6') || (MyReceiveBuffer[i] == '7') || (MyReceiveBuffer[i] == '8') || (MyReceiveBuffer[i] == '9') || (MyReceiveBuffer[i] == ':') || (MyReceiveBuffer[i] == '.') || (MyReceiveBuffer[i] == '/') || (MyReceiveBuffer[i] == ' '))// : . / "" 
						{
							dataascii = ascii.GetString(MyReceiveBuffer, i, 1);//string dataascii : on decompte 1 octet a la fois //inutile ?
							
							unelignedonnees += ascii.GetString(MyReceiveBuffer, i, 1);//probleme
							//tablignedonnees = unelignedonnees.ToCharArray();//

							string[] subs = unelignedonnees.Split(' ','	');//espace et tab
							/*foreach (var s in subs) //vois si il y a une autre methode pour avoir la taille du tableau
							{
								cpts++;
							}*/

							cpts=subs.Length;

							if (cpts==1) {
								string[] vdate = subs[0].Split('/');
								if ((vdate.Length == 3) && ((vdate[0].Length==1)|| (vdate[0].Length == 2)))
								{
									//textBox2.AppendText("Date :");//debug
									//textBox2.AppendText(subs[0]);
									strDate = subs[0];
								}
								else if ((vdate[0].Length > 2)||(vdate.Length > 3))
                                {
									filtredonnee = false;
                                }
							}

							if (cpts == 3)//il ya 2 espace entre la date et l'heure
							{
								//textBox2.AppendText("Heure :");
								//textBox2.AppendText(subs[2]); 
								strHeure = subs[2];

							}
							if (cpts == 5) {//il y a 2 espaces entre la date et la temperature
								//textBox2.AppendText("Temperature :");
								//textBox2.AppendText(subs[4]);
								strTemperature = subs[4];
							}
							if (cpts == 6) //il ya 2 espace
							{
								//if ((subs[5].Length == 2)|| (subs[5].Length == 3))//IL Y A UN PB ICI !!!
								//{
									//textBox2.AppendText("Lambda1 :");
									//textBox2.AppendText(subs[5]);
									strLambda1 = subs[5];
							//	}
								//else
                                //{
									//filtredonnee = false;
							//	}
							}
							if (cpts == 7)
							{
								
								if ((subs[6].Length <= 3))
								{
									//textBox2.AppendText("CHL :");
									//textBox2.AppendText(subs[6]);
									strCHL = subs[6];
								}
								else if (subs[6].Length > 3)
								{
									//filtredonnee = false;
								}
							}
							if (cpts == 8)
							{
								//textBox2.AppendText("Lambda2 :");
								//textBox2.AppendText(subs[7]);
								strLambda2 = subs[7];
							}
							if (cpts == 9)
							{
								//textBox2.AppendText("NTU :");
								//textBox2.AppendText(subs[8]);
								strNTU = subs[8];
							}
							if (cpts == 10)
							{
								//textBox2.AppendText("Thermistance :");
								//textBox2.AppendText(subs[9]);
								strThermistance = subs[9];
							}


							if (MyReceiveBuffer[i] == 13) {


								if ((filtredonnee == false) || (cpts != 10))//si la date est mauvaise et si on a pas toutes les donnes
								{
									filtredonnee = true;
								}
								else
								{
									r++;
									if (j == 0)
									{
										workRow = DatatableFLUNTR1.NewRow();
										workRow[1] = strDate;
										workRow[2] = strHeure;
										workRow[3] = strTemperature;
										workRow[4] = strLambda1;
										workRow[5] = strCHL;
										workRow[6] = strLambda2;
										workRow[7] = strNTU;
										workRow[8] = strThermistance;
										DatatableFLUNTR1.Rows.Add(workRow);
									}
									else if(j == 1)
									{
										workRow = DatatableFLUNTR2.NewRow();
										workRow[1] = strDate;
										workRow[2] = strHeure;
										workRow[3] = strTemperature;
										workRow[4] = strLambda1;
										workRow[5] = strCHL;
										workRow[6] = strLambda2;
										workRow[7] = strNTU;
										workRow[8] = strThermistance;
										DatatableFLUNTR2.Rows.Add(workRow);
									}
									else if(j == 2)//apres la trame 2  : trame 3
									{
										workRow = DatatableFLUNTR3.NewRow();
										workRow[1] = strDate;
										workRow[2] = strHeure;
										workRow[3] = strTemperature;
										workRow[4] = strLambda1;
										workRow[5] = strCHL;
										workRow[6] = strLambda2;
										workRow[7] = strNTU;
										workRow[8] = strThermistance;
										DatatableFLUNTR3.Rows.Add(workRow);
									}


									//dataGridView1.DataSource = DatatableFLUNTR; //ne maaaaarche pas

									/*
									textBox2.AppendText(newLine+"Dans la Datatable on a :"); //pour tester les donnes dans la database (pb !!!)
									for (int l = 0; l < 9; l++)
									{
										textBox2.AppendText(newLine +l+">"+ workRow[l]);
									}
									*/

									///////////////////////////////////////
									textBox2.AppendText(newLine);
									textBox2.AppendText(unelignedonnees);//on affiche les données correctes recu
									/*
									string[] Heureinterface = strHeure.Split(':');

									if (strHeure.Length == 3)
									{
										heuresbouee = Heureinterface[0];
										minutesbouee = Heureinterface[1];
										secondesbouee = Heureinterface[2];
									}
									*/
								}//end else //Acolade a remettre avec le else !!!!!!!

									cptligne++;
									cpts = 0;

									if (j == 0) { dataascii1 += unelignedonnees+newLine; }//on enregistre de cette facon les lignes de la trame 1 //pour plus tard on verifiera les checksum avant
									if (j == 1) { dataascii2 += unelignedonnees+newLine; }//trame 2
									if (j == 2) { dataascii3 += unelignedonnees+newLine; }//trame 3

									unelignedonnees = "";//on l'efface
								
							}//end if retour a la ligne
							
								//textBox3.Text = textBox3.Text + MyReceiveBuffer[i];//affiche val string de l'octet
								//sum = sum + MyReceiveBuffer[i]; //on va utiliser une fonction

						}

						/*if (MyReceiveBuffer[i] == 10) si jamais le programme compte les LF
						{ cptligne++; }*/


					}
					if ((MyReceiveBuffer[i] == '<') &&(oldcara=='<')&&(oldoldcara=='<')){ 
						FILTRE = false;
						testHorloge = false;
						//textBox2.Text = textBox2.Text + "   <   ";
					}

					////////* filtrer checksum envoyer
					if(testHorloge == true) {/////////////////////////////test de l'horloge


						unelignedonnees += ascii.GetString(MyReceiveBuffer, i, 1);
						string[] subs = unelignedonnees.Split(' ', ' ');//espace et tab
						{
							cpts++;
						}
						
					   cpts = subs.Length;
						
						if (cpts == 3)//il ya 2 espace entre la date et l'heure
						{
							 Heureinterface = subs[2].Split(':');
							//textBox2.AppendText("Heure :");
							//textBox2.AppendText(Heureinterface.Length.ToString());
							strHeure = subs[2];
						}
						if ((cpts>=3)&&(Heureinterface.Length == 3))
						{
							heuresbouee = Heureinterface[0];
							minutesbouee = Heureinterface[1];
							secondesbouee = Heureinterface[2];

							textBox2.AppendText(heuresbouee+" "+ minutesbouee);

							if ((heuresbouee!=DateTime.Now.Hour.ToString()) ||(minutesbouee!= DateTime.Now.Minute.ToString()))
                            {
								serialPort1.Write("g");
							}
							testHorloge = false;
						}
					}
					////////////

					if ((MyReceiveBuffer[i] == '*') && (FILTRE!=false))//A chaque *
					{
						TESTk = !TESTk;//pour le 1er Testk devient "false"
						//textBox2.Text = textBox2.Text + "   *  ";
						k++;
						if (k == 2)
						{
							valCS1 = checksumfiltrer;
							textBox3.AppendText(newLine + "CS1 = " + valCS1);
							checksumfiltrer = "";
						}
						if (k == 4)
						{
							valCS2 = checksumfiltrer;
							textBox3.AppendText(newLine + "CS2 = " + valCS2);
							checksumfiltrer = "";
						}
						if (k == 6)
						{
							valCS3 = checksumfiltrer;
							textBox3.AppendText(newLine + "CS3 = " + valCS3);
							checksumfiltrer = "";
							//k = 0;
						}
					}
					
					if ((TESTk == false) && (MyReceiveBuffer[i] != '*'))
					{
						dataascii = ascii.GetString(MyReceiveBuffer, i, 1);
						//textBox2.AppendText(dataascii);//a tester						 
						checksumfiltrer += dataascii;//on obtient val checksum  
					}
					/////////////
					if ((MyReceiveBuffer[i] == '#')&&(FILTRE != false))//A chaque #
					{
						TESTl = !TESTl;//pour le 1er Testk devient "false"
									   //textBox2.Text = textBox2.Text + "   #  ";
						l++;
						if (l == 2)
						{
							valnbrL1 = nbrlignefiltrer;
							textBox3.AppendText(newLine + "L1 = " + valnbrL1);
							nbrlignefiltrer = "";
						}
						if (l == 4)
						{
							valnbrL2 = nbrlignefiltrer;
							textBox3.AppendText(newLine + "L2 = " + valnbrL2);
							nbrlignefiltrer = "";
						}
						if (l == 6)
						{
							valnbrL3 = nbrlignefiltrer;
							textBox3.AppendText(newLine + "L3 = " + valnbrL3);
							nbrlignefiltrer = "";
							l = 0;
						}
					}
					if ((TESTl == false) && (MyReceiveBuffer[i] != '#'))
					{
						dataascii = ascii.GetString(MyReceiveBuffer, i, 1);
						//textBox2.AppendText(dataascii);//a tester						 
						nbrlignefiltrer += dataascii;//on obtient nbr ligne
					}

					_3oldcara = oldoldcara;
					oldoldcara = oldcara;
					oldcara = MyReceiveBuffer[i];//on garde en memoire l'ancien cara lue pour pouvoir le reutiliser
					
				}//end for

                }


			if (k == 6)//checksum
			{
				

				bool receptionok = false;
				//Au final on affiche une des trames de données exacte : 
				//textBox2.AppendText(somme1.ToString() + "val" + valCS1);
				//pb l'ecrit plusieur fois
				if (somme1.ToString() == valCS1)//avec verificationchecksomme=true; verificationnbrligne = true; si on veut verifier qu'au moins 2 trames sont identiques
				{
					textBox9.AppendText(newLine + dataascii1);
					serialPort1.Write("v");//donnees recu et verifiees
					numerotramevrai = 1;
					 receptionok =true;
					
				}
				else if (somme2.ToString() == valCS2)//avec verificationchecksomme=true; verificationnbrligne = true; si on veut verifier qu'au moins 2 trames sont identiques
				{
					textBox9.AppendText(newLine + dataascii2);
					serialPort1.Write("v");
					numerotramevrai = 2;
					receptionok = true;
				}
				else if (somme3.ToString() == valCS3)//avec verificationchecksomme=true; verificationnbrligne = true; si on veut verifier qu'au moins 2 trames sont identiques
				{
					textBox9.AppendText(newLine + dataascii3);
					serialPort1.Write("v");
					numerotramevrai = 3;
					receptionok = true;
				}
				
				if ((receptionok == true)&&(modeauto == true)) {//en mode auto seulement on ecrit toutes les data dans excel
					Thread t = new Thread(new ThreadStart(EcrituredataThread));//on va ecrire la ligne de donnée dans le excel ou gridview
					t.IsBackground = true;
					t.Start();
					joursansreception = 0;
					serialPort1.Write("z");
				}

				k = 0;

			}
				////////////////////verification affichage///////////////////// debug :

				verificationchecksomme =false;
				 verificationnbrligne=false;

				textBox7.Clear();
				textBox7.AppendText(newLine + "Calculer :"+ newLine + "  Checksome :"+ newLine + "1 : " + somme1 + newLine + "2 : " + somme2 + newLine + "3 : " + somme3);
				textBox7.AppendText(newLine + "   Nbr de lignes :" + newLine + "1 : " + nbrligne1 + newLine + "2 : " + nbrligne2 + newLine + "3 : " + nbrligne3);
				textBox7.AppendText(newLine + "Recu/Filtrer :" + newLine + "  Checksome :" + newLine + "1 : " + valCS1 + newLine + "2 : " + valCS2 + newLine + "3 : " + valCS3);
				textBox7.AppendText(newLine + "   Nbr de lignes :" + newLine + "1 : " + valnbrL1 + newLine + "2 : " + valnbrL2 + newLine + "3 : " + valnbrL3);
				textBox6.Clear();
				///verification des calculs des donnes recus


				if ((somme1 !=somme2) || (somme2 != somme3) || (somme1!=somme3) )
				{ textBox6.AppendText(newLine + " ! Les checksommes calculer sont differents "); 
					if ((somme1 == somme2) || (somme1 == somme3)) { checksommecalculervrai = somme1;
						verificationchecksomme = true;
					}
					if ((somme2 == somme1) || (somme2 == somme3)) { checksommecalculervrai = somme2;
						verificationchecksomme = true;
					}
					if ((somme3 == somme2) || (somme3 == somme1)) { checksommecalculervrai = somme3;
						verificationchecksomme = true;
					}
					if ((somme1 != somme2)&& (somme2 !=somme3)&& (somme1!=somme3)) { textBox6.AppendText(newLine + "!!! Aucuns checksommes identiques"); };
				}
				else { textBox6.AppendText(newLine + "Les 3 checksommes calculer sont identiques ! :" + somme1);
					verificationchecksomme = true;
					checksommecalculervrai = somme1;
				}
				if ((nbrligne1 != nbrligne2) || (nbrligne2 != nbrligne3) || (nbrligne1 != nbrligne3))
				{ textBox6.AppendText(newLine + " ! Les nbrs de lignes calculer sont differents "); 
					if ((nbrligne1 == nbrligne2) || (nbrligne1 == nbrligne3)) { nbrlignecalculervrai = nbrligne1;
						verificationnbrligne = true;
					}
					if ((nbrligne2 == nbrligne1) || (nbrligne2 == nbrligne3)) { nbrlignecalculervrai = nbrligne2;
						verificationnbrligne = true;
					}
					if ((nbrligne3 == nbrligne2) || (nbrligne3 == nbrligne1)) { nbrlignecalculervrai = nbrligne3;
						verificationnbrligne = true;
					}
					if ((nbrligne1 != nbrligne2) && (nbrligne2 != nbrligne3) && (nbrligne1 != nbrligne3)) { textBox6.AppendText(newLine + "!!! Aucuns nombres de lignes identiques"); }
				}
				else { textBox6.AppendText(newLine + "Les 3 nbrs de lignes calculer sont identiques ! :" + nbrligne1);
					if (nbrligne1 != 0)
					{
						verificationnbrligne = true;
						nbrlignecalculervrai = nbrligne1;
					}
				}


				///verification des données filtrer
				if ((valCS1 != valCS2) || (valCS2 != valCS3) || (valCS1 != valCS3))
				{ textBox6.AppendText(newLine + " ! Les checksommes recues sont differents ");
					if ((valCS1 == valCS2)  || (valCS1 == valCS3)) { checksommerecuvrai = valCS1; 
						if (verificationchecksomme == true)
						{
							if (checksommerecuvrai == checksommecalculervrai.ToString())
							{
								textBox6.AppendText(newLine + "Checksommes recus et verifier : " + checksommerecuvrai);
								//textBox9.AppendText(newLine + "Trame1 :" + dataascii1);
								label8.BackColor = Color.Green;
							}
						}
					}
					else if ((valCS2 == valCS1) || (valCS2 == valCS3)) { checksommerecuvrai = valCS2; 
						if (verificationchecksomme == true)
						{
							if (checksommerecuvrai == checksommecalculervrai.ToString())
							{
								textBox6.AppendText(newLine + "Checksommes recus et verifier : " + checksommerecuvrai);
								//textBox9.AppendText(newLine + "Trame2 :" + dataascii2);
								label8.BackColor = Color.Green;
							}
						}
					}
					else if ((valCS3 == valCS2) || (valCS3 == valCS1)) { checksommerecuvrai = valCS3; 
						if (verificationchecksomme == true)
						{
							if (checksommerecuvrai == checksommecalculervrai.ToString())
							{
								textBox6.AppendText(newLine + "Checksommes recus et verifier : " + checksommerecuvrai);
								//textBox9.AppendText(newLine + "Trame3 :" + dataascii3);
								label8.BackColor = Color.Green;
							}
						}
					}
					if ((valCS1 != valCS2) && (valCS2 != valCS3) && (valCS1 != valCS3)) { textBox6.AppendText(newLine + "!!! Aucuns checksums recus identiques"); }
				}
				else { textBox6.AppendText(newLine + "Les 3 checksommes recues sont identiques !"+ valCS1);
					checksommerecuvrai = valCS1;
					if (verificationchecksomme == true) {
						if (checksommerecuvrai == checksommecalculervrai.ToString()) {
							textBox6.AppendText(newLine + "Checksommes recus et verifier : " + checksommerecuvrai);
							//textBox9.AppendText(newLine + "Trame1 :" + dataascii1+ newLine + "Trame2 :" + dataascii2 +newLine + "Trame3 :" + dataascii3);
							label8.BackColor = Color.Green;
						}
						}
				}
				if (verificationchecksomme==false) { label8.BackColor = Color.Red ; }
				if ((valnbrL1 != valnbrL2) || (valnbrL2 != valnbrL3) || (valnbrL1 != valnbrL3))
				{ textBox6.AppendText(newLine + " ! Les nbrs de lignes recus sont differents ");
					if ((valnbrL1 == valnbrL2) || (valnbrL1 == valnbrL3)) { nbrlignerecuvrai = valnbrL1;
						if (verificationnbrligne == true)
						{
							if (nbrlignerecuvrai == nbrlignecalculervrai.ToString())
							{
								textBox6.AppendText(newLine + "Nbr de lignes recus et verifier : " + nbrlignecalculervrai);//probleme l'affiche 2 fois 
							}
						}
					}
					else if ((valnbrL2 == valnbrL1) || (valnbrL2 == valnbrL3)) { nbrlignerecuvrai = valnbrL2;
						if (verificationnbrligne == true)
						{
							if (nbrlignerecuvrai == nbrlignecalculervrai.ToString())
							{
								textBox6.AppendText(newLine + "Nbr de lignes recus et verifier : " + nbrlignecalculervrai);
							}
						}
					}
					else if ((valnbrL3 == valnbrL2) || (valnbrL3 == valnbrL1)) { nbrlignerecuvrai = valnbrL3;
						if (verificationnbrligne == true)
						{
							if (nbrlignerecuvrai == nbrlignecalculervrai.ToString())
							{
								textBox6.AppendText(newLine + "Nbr de lignes recus et verifier : " + nbrlignecalculervrai);
							}
						}
					}
					if ((valnbrL1 != valnbrL2) && (valnbrL2 != valnbrL3) && (valnbrL1 != valnbrL3)) { textBox6.AppendText(newLine + "!!! Aucuns checksums recus identiques"); }
				}
				else { textBox6.AppendText(newLine + "Les 3 nbrs de lignes recus sont identiques ! :" + valnbrL1);
					nbrlignerecuvrai = valnbrL1;
					if (verificationnbrligne == true)
					{
						if (nbrlignerecuvrai == nbrlignecalculervrai.ToString())
						{
							textBox6.AppendText(newLine + "Nbr de lignes recus et verifier : " + nbrlignecalculervrai);
						}
					}
				}//end else

				//prevoir que faire si aucuns des CS marchent

				//byte[] dataasciirecu = dataascii; //?
				/* //en utilisant un tableau de char a la place :///////////////////////////////////////////////////////
				for (int i=0; i<tab.Length-1;i++) //
				{ 
					textBox2.Text = textBox2.Text + "+" + tab[i];//en utilisant un tableau de char
					if (tab[i] == '>')
					{
						TEST = true;
						textBox2.Text = textBox2.Text + ">   ";
					}
					if (tab[i] == '<')
					{
						TEST = false;
						textBox2.Text = textBox2.Text + "   <";
					}
					if (TEST == false)
					{
						//tabdata[j] = tab[i];
						//j++;
					}
				}
				*/
				/*if (tabdata != null)
				{
					textBox2.Text = tabdata.ToString();
				}*/
			
			/*if (myReceiveData != null)//FAUX
			{
				byte sumdata = CheckSum(myReceiveData);
			}
			textBox3.Text = sumdata.ToString(); 
			*/

		}//end serialport1_DataReceived
		private static int CheckSum(byte byteData) //byte c(byte[] ) //pour l'instant on a qu'1 cara a la fois 
		{
			int chkSumByte = 0; //byte = 0x00
								//for (int i = 0; i < byteData.Length; i++) //s'il y a plusieurs cara
								//chkSumByte ^= byteData[i];
			chkSumByte += byteData;
			return chkSumByte;
		}

		private void modeAUTO() //Activer si on a activer le mode auto
        {

		/*	textBox1.AppendText(newLine + "(Envoie Test)");
			serialPort1.Write("t");
			ReceptionAuto = true;
		*/
			//D'abord on envoie t pour verifier l'heure (on recoit test 25 / 5 / 2022  13:55:42  28.00 recu)
			textBox1.AppendText(newLine + "(Demande de données automatique)");
			serialPort1.Write("l");
			joursansreception++; //ce jour si on a voulu recuperer des donnes
		}
		private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
		{
			if (serialPort1.PortName != comboBox1.SelectedItem.ToString())//si on change de com
			{
				// Fermeture du port COM actuel
				if (serialPort1.IsOpen == true)
					serialPort1.Close();
				// Prise en compte du port COM choisi par l'utilisateur
				serialPort1.PortName = comboBox1.SelectedItem.ToString();
				// Adaptation du texte du bouton de commande
				//button1.Text = "OK" + serialPort1.PortName;
				// Ouverture du nouveau port COM
				if (serialPort1.IsOpen == false)
					try { serialPort1.Open(); }
					catch ( Exception e3)
                    {
						MessageBox.Show(e3.Message, "Erreur !");
                    }
			}
		}
		private void button3_Click(object sender, EventArgs e)
		{
			serialPort1.Write("t");
			testHorloge = true;
		}
		private void button4_Click(object sender, EventArgs e)
		{
			j = 0;//on reintialise le cpt qui compte le nbr de trame de donnes <"..."> envoyer
			serialPort1.Write("l");	
		}
		private void button5_Click(object sender, EventArgs e)
		{
			serialPort1.Write("c");
		}
		private void button9_Click(object sender, EventArgs e)
		{
			serialPort1.Write("g");
		}
        private void button6_Click(object sender, EventArgs e)
        {
			textBox1.Clear();
        }
        private void button7_Click(object sender, EventArgs e)
        {
			textBox2.Clear();
        }

		private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)//si on change le baud rate
        {
			///if (serialPort1.IsOpen == true)
			///serialPort1.Close();
			serialPort1.DiscardInBuffer();//ou supprime les donnes encore dans le buffer pour eviter d'avoir des donnes fausse
			int newbaudrate = Convert.ToInt32(comboBox2.SelectedItem);
			if (newbaudrate > 0)
			{
				serialPort1.BaudRate = newbaudrate; 
				textBox8.Text=(serialPort1.BaudRate).ToString();
			}
			///if (serialPort1.IsOpen == false)
			///serialPort1.Open();
		}
		private void button2_Click(object sender, EventArgs e) //button "mode auto"
		{
			modeauto = true;
			button2.BackColor = Color.Green;
			button8.BackColor = Color.Transparent;
		}

		private void button8_Click(object sender, EventArgs e) //mode manuel
        {
			modeauto = false;
			button8.BackColor = Color.Green;
			button2.BackColor = Color.Transparent;
		}

		private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
		{
			compteurc = textBox4.Text;
			if (compteurc.Length == 4)
			{
				serialPort1.Write(compteurc + "\n");
				textBox4.Clear();
				textBox5.Text = "Vous avez entrer : " + compteurc;
			}
		}

		private void textBox11_KeyPress(object sender, KeyPressEventArgs e)//pour le saisir le nbr d'acquisition ("00 à 99");
		{
			cptaqui = textBox11.Text;
			if (cptaqui.Length == 2)
			{
				serialPort1.Write(cptaqui + "\n");
				textBox11.Clear();
				textBox5.Text = "Vous avez entrer : " + cptaqui;
			}
		}
		private void button10_Click(object sender, EventArgs e)//demande nbr d'aquisition
		{
			serialPort1.Write("n");
		}

		private void button11_Click(object sender, EventArgs e)//RESET
		{
			serialPort1.Write("r");
		}

        private void button12_Click(object sender, EventArgs e)//Test database ou test datable aussi
        {
			try
			{
				String pathfile = @datatxtpathfile;
				Thread t = new Thread(new ThreadStart(EcrituredataThread));//on va ecrire la ligne de donnée dans le excel ou gridview
				t.IsBackground = true;
				t.Start();
				MessageBox.Show("Data has been saved to" + pathfile, "savefile");
			}
			        
				catch (Exception e3)
			{
				MessageBox.Show(e3.Message, "error");
			}
		
		}
		// // // // // // // // // // // // // // // // 
		// // // // // // // // // // // // // // // /
		private void EcrituredataThread()
        {
			Invoke(new ecrituredataDelegate(ecrituredata));
        }
		private void ecrituredata()
		{
			if ((numerotramevrai == 1) || (numerotramevrai == 2) || (numerotramevrai == 3))
			{
				if (modeauto == true)
				{
					//if (textBox9.Text =="")//s'il n'y a pas de texte
					//{
					//textBox1.AppendText(newLine+DateTime.Now.Day + DateTime.Now.Hour + DateTime.Now.Minute + ": Rien n'a été recu");
					//
					//}
					String pathfile = @datatxtpathfile;
					//String pathfile = @"C:\Users\grisoni\Desktop\";// @"C:\Users\ROMAIN.000\Desktop\"
					String filename = "data_FLUNTR.txt";
					System.IO.File.AppendAllText(pathfile + filename, textBox9.Text);//textbox2>>>9//on sauvegarde les données filtrer //a remplacer par la text box des donnes filtrer verifier quand le checksum sera OK
					textBox9.Clear();
				}
			}

			//mettre code pour l'ecriture dans datagrid
			//string[] Rows = { strDate, strHeure, strTemperature, strLambda1, strCHL, strLambda2, strNTU, strThermistance };
			//dataGridView1.Rows.Add(Rows);//marche mais trop lent dc on le met pas
			//dataGridView1.AutoResizeRows();

			 //1
			Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Open(@datatxtpathfile + "output.xlsx");
			Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
			app.Visible = true;
			worksheet = workbook.Sheets[1];//sheet 1 ou feuil1
			worksheet = workbook.ActiveSheet;
			worksheet.Name = "Data";
			





			

			
			if (numerotramevrai == 1)
			{
				int m = 0;
				int n = 0;
				foreach (DataRow row in DatatableFLUNTR1.Rows)
				{
					m++;
					foreach (DataColumn column in DatatableFLUNTR1.Columns)
					{
						n++;
						String STdataFLUNTR = row[column].ToString();
						//textBox1.AppendText(STdataFLUNTR + " ");
						worksheet.Cells[m+ligneexcel, n] = STdataFLUNTR;//+ Ajouter la valeur de la derniére ligne
					}
					n = n - 9;
					//textBox1.AppendText(newLine);
				}
				ligneexcel += m;

				Properties.Settings.Default.nombredeligneexcel = ligneexcel;
				Properties.Settings.Default.Save();
			}

			else if (numerotramevrai == 2)
			{
				int m = 0;
				int n = 0;
				foreach (DataRow row in DatatableFLUNTR2.Rows)
				{
					m++;
					foreach (DataColumn column in DatatableFLUNTR2.Columns)
					{
						n++;
						String STdataFLUNTR = row[column].ToString();
						//textBox1.AppendText(STdataFLUNTR + " ");
						worksheet.Cells[m+ligneexcel, n] = STdataFLUNTR;//+ Ajouter la valeur de la derniére ligne
					}
					n = n - 9;
					//textBox1.AppendText(newLine);
				}
				ligneexcel += m;
				Properties.Settings.Default.nombredeligneexcel = ligneexcel;
				Properties.Settings.Default.Save();
			}
						
			else if (numerotramevrai == 3)
			{
				int m = 0;
				int n = 0;
				foreach (DataRow row in DatatableFLUNTR3.Rows)
				{
					m++;
					foreach (DataColumn column in DatatableFLUNTR3.Columns)
					{
						n++;
						String STdataFLUNTR = row[column].ToString();
						//textBox1.AppendText(STdataFLUNTR + " ");
						worksheet.Cells[m+ligneexcel, n] = STdataFLUNTR;//+ Ajouter la valeur de la derniére ligne
					}
					n = n - 9;

					//textBox1.AppendText(newLine);
				}
				ligneexcel += m;
				Properties.Settings.Default.nombredeligneexcel = ligneexcel;
				Properties.Settings.Default.Save();
			}
			

			//}
			
			/*
			var saveFileDialoge = new SaveFileDialog();
			saveFileDialoge.FileName = "output";
			saveFileDialoge.DefaultExt = ".xlsx";
			if (saveFileDialoge.ShowDialog() == DialogResult.OK)
			{
				workbook.SaveAs(saveFileDialoge.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
			}
			*/
			//app.Quit();

		}

        private void textBox13_KeyPress(object sender, KeyPressEventArgs e)
        {
			string cptcommande = textBox13.Text;
			if (cptcommande.Length == 1)
			{
				serialPort1.Write(cptcommande);
				textBox5.Text = "Vous avez entrer : " + cptcommande;
				textBox13.Clear();
			}
		}

        private void button13_Click(object sender, EventArgs e)
        {
			serialPort1.Write("z");
		}

        private void button14_Click(object sender, EventArgs e)
        {
			try
			{
				String pathfile = @datatxtpathfile;
				//String pathfile = @"C:\Users\grisoni\Desktop\";// @"C:\Users\ROMAIN.000\Desktop\"
				String filename = "data_FLUNTR.txt";
				System.IO.File.AppendAllText(pathfile + filename, textBox9.Text);//on sauvegarde les données filtrer//2>>9
				MessageBox.Show("Data has been saved to" + pathfile, "savefile");
				//dataasciivrai

			}
			catch (Exception e2)
			{
				MessageBox.Show(e2.Message, "error");
			}
		}

        private void button16_Click(object sender, EventArgs e)
        {
			@datatxtpathfile = textBox14.Text;

			Properties.Settings.Default.EmplacementFichier = @datatxtpathfile;
			Properties.Settings.Default.Save();
			label19.Text=@datatxtpathfile;
			label19.ForeColor = Color.Black;

        }

		private void initDatatableFLUNTR1()
        {
			DatatableFLUNTR1.Reset();
			DatatableFLUNTR1.Columns.Add("Table_ID", typeof(int));
			DatatableFLUNTR1.Columns.Add("Date", typeof(String));
			DatatableFLUNTR1.Columns.Add("Heure", typeof(String));
			DatatableFLUNTR1.Columns.Add("Temperature", typeof(String));
			DatatableFLUNTR1.Columns.Add("Lambda1", typeof(String));
			DatatableFLUNTR1.Columns.Add("CHL", typeof(String));
			DatatableFLUNTR1.Columns.Add("Lambda2", typeof(String));
			DatatableFLUNTR1.Columns.Add("NTU", typeof(String));
			DatatableFLUNTR1.Columns.Add("Thermistance", typeof(String));
		}
		private void initDatatableFLUNTR2()
		{
			DatatableFLUNTR2.Reset();
			DatatableFLUNTR2.Columns.Add("Table_ID", typeof(int));
			DatatableFLUNTR2.Columns.Add("Date", typeof(String));
			DatatableFLUNTR2.Columns.Add("Heure", typeof(String));
			DatatableFLUNTR2.Columns.Add("Temperature", typeof(String));
			DatatableFLUNTR2.Columns.Add("Lambda1", typeof(String));
			DatatableFLUNTR2.Columns.Add("CHL", typeof(String));
			DatatableFLUNTR2.Columns.Add("Lambda2", typeof(String));
			DatatableFLUNTR2.Columns.Add("NTU", typeof(String));
			DatatableFLUNTR2.Columns.Add("Thermistance", typeof(String));
		}
		private void initDatatableFLUNTR3()
		{
			DatatableFLUNTR3.Reset();
			DatatableFLUNTR3.Columns.Add("Table_ID", typeof(int));
			DatatableFLUNTR3.Columns.Add("Date", typeof(String));
			DatatableFLUNTR3.Columns.Add("Heure", typeof(String));
			DatatableFLUNTR3.Columns.Add("Temperature", typeof(String));
			DatatableFLUNTR3.Columns.Add("Lambda1", typeof(String));
			DatatableFLUNTR3.Columns.Add("CHL", typeof(String));
			DatatableFLUNTR3.Columns.Add("Lambda2", typeof(String));
			DatatableFLUNTR3.Columns.Add("NTU", typeof(String));
			DatatableFLUNTR3.Columns.Add("Thermistance", typeof(String));
		}

        private void button17_Click(object sender, EventArgs e)
        {
			textBox9.Clear();//on efface les donnes filtrervalider
        }

        private void button1_Click(object sender, EventArgs e)//aff fenetres pour debug les checksums et nbr lignes
        {
			panel4.Hide();
			button1.Hide();
		}

        private void button18_Click(object sender, EventArgs e)
        {
			panel4.Show();
			button1.Show();
		}

        private void button19_Click(object sender, EventArgs e)
        {
			panel3.Hide();
			button19.Hide();
		}

        private void button20_Click(object sender, EventArgs e)
        {

			panel3.Show();
			button19.Show();
		}

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

		}

        private void button21_Click(object sender, EventArgs e)
        {
			
			System.Diagnostics.Process.Start(Application.ExecutablePath);
			Application.Exit();//pb ?
		}

        private void button15_Click(object sender, EventArgs e)
        {
			ligneexcel=0;
        }

        private void button22_Click(object sender, EventArgs e)
        {
			ligneexcel = Properties.Settings.Default.nombredeligneexcel;

		}
    }

    public delegate void ecrituredataDelegate();
	

}
