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
using Excel = Microsoft.Office.Interop.Excel;//Cette espace de nom va nous permettre d'écrire des données dans un fichier excel (la version d'excel doit être au moins celle de 2007)

namespace App_1 
{
	public partial class Form1 : Form
	{
		DataTable DatatableFLUNTR1 = new DataTable();
		DataTable DatatableFLUNTR2 = new DataTable();
		DataTable DatatableFLUNTR3 = new DataTable();
		
		ASCIIEncoding ascii;//Alphabet de decodage
		
		int sum;//variable utiliser pour additionner les valeurs des caractéres = checksomme
		int cptligne = 0;//variable utiliser pour compter le nombre de ligne dans chaques trames
		int somme1;//Variable où l'on va ecrire le checksomme de la trame 1
		int somme2;//Variable où l'on va ecrire le checksomme de la trame 2
		int somme3;//Variable où l'on va ecrire le checksomme de la trame 3
		
		int nbrligne1;//Variable où l'on va ecrire le nombre de ligne de la trame 1
		int nbrligne2;//Variable où l'on va ecrire le nombre de ligne de la trame 2
		int nbrligne3;//Variable où l'on va ecrire le nombre de ligne de la trame 3
		
		int j;//Compte le nombre de trame, s'incrémente à chaque "<<<[...]>>>"
		int k;//Compte le nombre de fois que l'on recoit "*[...]*" :checksomme envoyer par l'interface
		int l;//Compte le nombre de fois que l'on recoit "#[...]#" :nombre de ligne envoyer par l'interface
		
		int checksommecalculervrai;//variable utiliser pour afficher le checksomme d'une trame "vérifié" > avec un checksomme identique que celui envoyer par l'interface à distance
		int nbrlignecalculervrai;//de même mais pour afficher le nombre de ligne
		int cpts = 0;//permet de compter le nombre d'espaces ou de string créer en découpant une chaine de caractére de données (découper par les espaces ou tab)
		int numerotramevrai;//enregsitre le nurmero d'une trame vérifié si au moins une des 3 trames de données entre "<<<>>>" envoyés par le capteur est "vrai" 
		int ligneexcel;//variable qui garde en memoire la derniére ligne où l'on a ecrit dans le fichier excel pour ecrire à sa suite
		int joursansreception;//garde en mémoire le nombre de jour où l'acquisition de donnée à échouer
		const int MAXDATA = 1048576;
		
		string newLine = Environment.NewLine;//L'utilisation de ce string permet de sauter une ligne
		string dataascii1;//string qui enregistre les données filtrés de la trame 1
		string dataascii2;//string qui enregistre les données filtrés de la trame 2
		string dataascii3;//string qui enregistre les données filtrés de la trame 3
		
		string checksumfiltrer;//chaine de caractére où en enregsitre les caractéres lu entre "# #"
		string valCS1;//varaiable où l'on enregistre le checksumfiltrer de la trame 1
		string valCS2;//idem mais pour la trame 2
		string valCS3;//idem mais pour la trame 3
		string valnbrL1;//chaine de caractére où l'on enregsitre la valeur du nombre de ligne filtré (entre ##) pour la trame1
		string valnbrL2;//idem pour la tramme 2
		string valnbrL3;//idem pour la 3
		
		string nbrlignefiltrer;//la valeur filtrer entre ##
		String compteurc;//compte le nombre de caractére dans la textbox du bouton "compteur", a 4 chiffre en enverra la nouvelle frequence d'acquisition de donnée
		String cptaqui;//de même pour la textbox associer au bouton "nbr d'acquisition"
		String dataascii;
		string checksommerecuvrai;//pour les fenetres de debug où on enregistre le checksomme si le checksomme filtré entre ** et égale à celui calculer par le programme
		string nbrlignerecuvrai;//de même avec le nombre de ligne/
		string unelignedonnees;//chaine de caractére qui enregistre une à la fois une ligne de données filtrés pendant la reception des trames de données.
		String strDate;//String obtenu aprés le découpage d'une ligne de donnée, correspond au paramétre de la date 
		String strHeure;//de même et correspond à l'heure de l'interface
		String strTemperature;//corespond à la température
		String strLambda1;//la valeur du paramétre lambda1 mesurer par le capteur 
		String strCHL;//la valeur du paramétre CHL mesurer par le capteur
		String strLambda2;//la valeur du paramétre lambda2 mesurer par le capteur
		String strNTU;//la valeur du paramétre NTU mesurer par le capteur
		String strThermistance;//la valeur du paramétre de "thermistance" mesurer par le capteur fluntr
		String @datatxtpathfile; //l'emplacement où les fichier sont enregistrer
		String heuresbouee;//string où l'on enregistre l'heure de la bouée, filtrer à partir du string strHeure
		String minutesbouee;//de même pour les minutes de la bouée
		String secondesbouee;//et les secondes //cette varaible n'est toutefois pas utiliser pour l'instant
		String[] Heureinterface;//tableau de char issue de la chaine de caractére de l'heure (=strHeure) qui sert à séparer l'heure, les minutes et secondes du string "strheure"
		
		byte oldcara;//enregistre l'ancien caractére lu par la boucle FOR des caractéres recu par liaison serie
		byte oldoldcara;//enregistre l'ancien oldcara
		byte _3oldcara;//enregistre l'ancien oldoldcara
		byte[] MyReceiveBuffer;//tableau de byte qui va contenir la totalité des caratéres recu par la liaison serie
		
		bool FILTRE = false;//Vrai si on filtre les données entre <<< >>>
		bool filtredonnee;//indique si les paramétres filtrer ou une ligne dde donnée est correcte /à un bon format
		bool TESTk = false;//Vrai Quand on filtre les caractéres entre **
		bool TESTl = false;//Vrai quand on filtre les caractéres entre ##
		bool modeauto = false;//mode auto desactiver par defaut
		bool verificationchecksomme = false;//pour le debug vrai si on a bien un bon checksomme
		bool verificationnbrligne = false;//de même si les nombres de lignes sont identiques
		bool testHorloge =false;//Vrai quand on vérifie l'heure de l'interface (en mode auto aprés avoir envoyer "t")
		
		DataRow workRow;//nous permet d'ecrire dans une ligne des datatables

		private static System.Threading.Timer time1;//timer utiliser pour obtenir et afficher l'heure du PC 

		Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();//permet d'utiliser excel l'application est appellée "app"

		public Form1() 
		{
			InitializeComponent();
			CheckForIllegalCrossThreadCalls = false;
			serialPort1.Open();
			textBox8.Text = (serialPort1.BaudRate).ToString();//
			MyReceiveBuffer = new byte[MAXDATA];
			ascii = new ASCIIEncoding();//permet de construire l'alphabet ascii
			serialPort1.Parity = Parity.None;//reglage des paramétres du port serie
			serialPort1.DataBits = 8;
			serialPort1.StopBits = StopBits.One;
			time1 = new System.Threading.Timer(Timenow, null, 0, 1000);//toutes les 1 sec appelle  datetime
			@datatxtpathfile = Properties.Settings.Default.EmplacementFichier;//on récupére l'emplacement du fichier si celui-ci était enregistrer dans les paramétres du logiciel
			label19.Text = @datatxtpathfile;
			label19.ForeColor = Color.Black;
			ligneexcel = Properties.Settings.Default.nombredeligneexcel;//on récupére aussi la derniere ligne ou on a ecrit dans le fichier excel
		}
		private void Form1_Load(object sender, EventArgs e)
		{
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
		public void Timenow(Object state)//Appellée chaque secondes
        {
			textBox10.Clear();
			textBox10.Text = DateTime.Now.ToString();//on affiche la date
			
			if (DateTime.Now.Hour == 13 && DateTime.Now.Minute == 31 && DateTime.Now.Second == 00)
			{
				if (modeauto == true)
				{
					initDatatableFLUNTR1();//On est alors un nouveau jour et on efface les textbox et les variables au cas ou celles-ci ne c'étaient pas remise à 0 
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
					serialPort1.Write("t");//on demande un test à l'interface à distance
					testHorloge = true;//on est donc pendant un test de l'horloge de l'interface a distance
				}
			}
			if (DateTime.Now.Hour ==13 && DateTime.Now.Minute==35 && DateTime.Now.Second==00) {//a 13:35 on va faire une acquisition de donnée
				if (modeauto==true) {
					if (joursansreception >= 5)//pour éviter de recevoir trop de donnée, si la transmission ne fonctionne pas pour au moins 5 J 
					{
						serialPort1.Write("v");//on demande à l'interface de placer le curseur de sa carte SD "en bas" du fichier de facon à ne pas recevoir les derniéres données non envoyés  
					}
					textBox1.AppendText(newLine + "(Demande de données automatique)");
					serialPort1.Write("l");//on fait une acquisition de donnée
					joursansreception++;//on retient avoir demander une acquisition de donnée ce jour-ci (la varaible reviendra à 0si l'acquisition réussi et si au moins 1 trames sont correctes)
				}
			}
		}
		private void serialPort1_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
		{
			
			String decodedAscii;
			int iNbByteReceived;
			
			iNbByteReceived = serialPort1.BytesToRead;
			
			if (iNbByteReceived != 0)//pour éviter des erreurs au cas où la methode se declenche même sans byte recu
			{
				
				// On place ces octets dans notre buffer
				
				serialPort1.Read(MyReceiveBuffer, 0, iNbByteReceived);// On sort du buffer de reception du driver serie le nombre d'octets reçus
				decodedAscii = ascii.GetString(MyReceiveBuffer, 0, iNbByteReceived);// On interprete le tableau d'octet comme une chaine de caractere ASCII
				// affichage de la chaine ascii dans textBox1
				textBox1.AppendText(decodedAscii);
				textBox1.SelectionStart = textBox1.Text.Length;
				textBox1.ScrollToCaret();
				textBox1.Refresh();

				for (int i = 0; i < iNbByteReceived; i++){//on va traiter 1 par 1 chaques caractéres recus par la liaison serie
					if (testHorloge == true)//si on est pendant le test de l'horloge
					{
						unelignedonnees += ascii.GetString(MyReceiveBuffer, i, 1);
						string[] subs = unelignedonnees.Split(' ', ' ');//on decoupe la chaine de caractére par ses espace et tab
						
						cpts = subs.Length;//et on compte le nombre de dcoupage realiser/ de string différent obtenu

						if (cpts == 3)//il ya 2 espace entre la date et l'heure
						{
							Heureinterface = subs[2].Split(':');//on decoupe l'heure pour obetnir les secondes,minutes et heures separés
							strHeure = subs[2];
						}
						if ((cpts >= 3) && (Heureinterface.Length == 3))
						{
							heuresbouee = Heureinterface[0];
							minutesbouee = Heureinterface[1];
							secondesbouee = Heureinterface[2];

							textBox2.AppendText(heuresbouee + " " + minutesbouee);

							if ((heuresbouee != DateTime.Now.Hour.ToString()) || (minutesbouee != DateTime.Now.Minute.ToString()))
							{
								serialPort1.Write("g");//on demane au GPS de l'interface de remmettre à l'heure celle-ci
							}
							testHorloge = false;//fin du test horloge
						}
						if (MyReceiveBuffer[i] == 13)
                        {
							unelignedonnees = "";//a chaque saut de ligne on traite une nouvelle ligne
						}
					}

					if ((MyReceiveBuffer[i] == '>') && (oldcara == '>') && (oldoldcara == '>')) // à la fin de la récéption du trame de donnée
					{
						FILTRE = false;//on ne filtre plus les caractéres entre <<<>>>
						j++;
						if (j == 1)//si on est la 1ére trame
						{
							somme1 = sum; //on enregistre la valeur du checksomme calculer pour cette trame dans une variable
							textBox3.AppendText(newLine + "S1 = " + somme1.ToString());
							sum = 0;
							nbrligne1 = cptligne;//on enregistre le nombre de lignes calculer
							cptligne = 0;
						}
						if (j == 2)//de même pour la 2éme trame
						{
							somme2 = sum;
							textBox3.AppendText(newLine + "S2 = " + somme2.ToString());
							sum = 0;
							nbrligne2 = cptligne;
							cptligne = 0;
						}
						if (j == 3)//de même pour la 3éme
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
					

					if ((FILTRE == true )&&(MyReceiveBuffer[i]!='>'))// quand on est apres <<< avant >>> sans compter les ">>"
					{
						sum += MyReceiveBuffer[i];//on additionne la valeur en decimal des caractéres ascii
						if (oldcara == '>')//étant donnée que MyReceiveBuffer[i]='>' n'est pas pris en compte dans le calcul du checksomme, pour éviter de compter les >>> d'une fin de trame, lire le caractére précédent permet de détecter les occasionnels ">" qui serait apparu avec des données fausse/corrompus au millieu de la trame
						{ sum += 62; }
						if ((oldcara == '>') && (oldoldcara == '>'))//de même si 2 ">>" se suivent (bien que nous n'avons jamais observé ce cas)
						{ sum += 62; }

						if ( (MyReceiveBuffer[i]==13) ||(MyReceiveBuffer[i] =='	') || (MyReceiveBuffer[i] == '0') || (MyReceiveBuffer[i] == '1') || (MyReceiveBuffer[i] == '2') || (MyReceiveBuffer[i] == '3') || (MyReceiveBuffer[i] == '4') || (MyReceiveBuffer[i] == '5') || (MyReceiveBuffer[i] == '6') || (MyReceiveBuffer[i] == '7') || (MyReceiveBuffer[i] == '8') || (MyReceiveBuffer[i] == '9') || (MyReceiveBuffer[i] == ':') || (MyReceiveBuffer[i] == '.') || (MyReceiveBuffer[i] == '/') || (MyReceiveBuffer[i] == ' '))// : . / "" 
						{//On filtres les données fausses/corrompus pour ne garder que les caracétres qui constitue une ligne de donnée vrai
							dataascii = ascii.GetString(MyReceiveBuffer, i, 1);//string dataascii : on decompte 1 octet a la fois //inutile ?
							unelignedonnees += ascii.GetString(MyReceiveBuffer, i, 1);//probleme
							//tablignedonnees = unelignedonnees.ToCharArray();//

							string[] subs = unelignedonnees.Split(' ','	');//là encore on découpe la ligne de donnée par ses espaces et tabulations
							cpts=subs.Length;
							if (cpts==1) {
								string[] vdate = subs[0].Split('/');//on isole le jour,mois et année
								if ((vdate.Length == 3) && ((vdate[0].Length==1)|| (vdate[0].Length == 2)))
								{
									strDate = subs[0];
								}
								else if ((vdate[0].Length > 2)||(vdate.Length > 3))//certaines lignes de données corrompus commencent par 99/99/99/ ou similaire par exemple
                                {
									filtredonnee = false;
                                }
							}

							if (cpts == 3)//il ya 2 espace entre la date et l'heure, voilà pourquoi c'est le cpts 3 et non le 2éme
							{
								strHeure = subs[2];
							}
							if (cpts == 5) {
								strTemperature = subs[4];
							}
							if (cpts == 6)
                            { 
									strLambda1 = subs[5];
							}
							if (cpts == 7)
							{
								
								if ((subs[6].Length <= 3))
								{
									strCHL = subs[6];
								}
								else if (subs[6].Length > 3)//si une ligne de donnée est fausse il est fréquent que ce soit à partir du paramétre 5 celui du CHL
								{
								}
							}
							if (cpts == 8)
							{
								strLambda2 = subs[7];
							}
							if (cpts == 9)
							{
								strNTU = subs[8];
							}
							if (cpts == 10)
							{
								strThermistance = subs[9];
							}


							if (MyReceiveBuffer[i] == 13) {


								if ((filtredonnee == false) || (cpts != 10))//si la date est mauvaise et si on a pas toutes les donnes
								{
									filtredonnee = true;
								}
								else
								{
									if (j == 0)//si la ligne de donnée est vrai on la garde en méoire dans une datable avec les autres lignes précédentes de la trame
									{//on en profite pour isoler dès maintenant les paramétres dans différentes cases pour plus tard les écrire dans des cases séparés sur excel
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
									else if(j == 2) //  : trame 3
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

									textBox2.AppendText(newLine);//on écrit ces lignes dans une des boite de debug
									textBox2.AppendText(unelignedonnees);

								}   //end else 

									cptligne++;
									cpts = 0;

									if (j == 0) { dataascii1 += unelignedonnees+newLine; }//on enregistre de cette facon les lignes de la trame 1 //pour plus tard on verifiera les checksum avant
									if (j == 1) { dataascii2 += unelignedonnees+newLine; }//trame 2
									if (j == 2) { dataascii3 += unelignedonnees+newLine; }//trame 3

									unelignedonnees = "";//on l'efface
								
							}//end if retour a la ligne

						}

					}
					if ((MyReceiveBuffer[i] == '<') &&(oldcara=='<')&&(oldoldcara=='<')){ //Quand on débute la lecture d'une trame
						FILTRE = true;//On autorise le filtrage entre <<< >>>
						testHorloge = false;//Si jamais le testHorloge ne c'était pas arréter 
						
					}

					if ((MyReceiveBuffer[i] == '*') && (FILTRE!=true))//A chaque *
					{
						TESTk = !TESTk;//Au premier * on lit les valeurs qui suivent puis au 2nd on arréte on aura bien des trames : trame1 *valeur* trame2 *valeur* trame3 *valeur* 
						k++;
						if (k == 2) //C'est à dire que l'on a recu *<valeur>* qui corresponds au checksomme calculer par l'interface sur la bouée
						{
							valCS1 = checksumfiltrer;//on enregistre cette <valeur>
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
							k = 0;
						}
					}
					
					if ((TESTk == true) && (MyReceiveBuffer[i] != '*'))
					{
						dataascii = ascii.GetString(MyReceiveBuffer, i, 1);				 
						checksumfiltrer += dataascii;//on obtient la valeur du checksum  
					}
					
					if ((MyReceiveBuffer[i] == '#')&&(FILTRE != true))//A chaque #
					{
						TESTl = !TESTl;//même principe que pour le filtrage des valeurs entre **
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
					if ((TESTl == true) && (MyReceiveBuffer[i] != '#'))
					{
						dataascii = ascii.GetString(MyReceiveBuffer, i, 1);		 
						nbrlignefiltrer += dataascii;//on obtient nbr ligne
					}

					_3oldcara = oldoldcara;
					oldoldcara = oldcara;//on garde en mémoire l'ancien oldacara 
					oldcara = MyReceiveBuffer[i];//on garde en memoire l'ancien cara lue pour pouvoir le reutiliser
					
				}//end for

                }// endif iNbBytereceived!=0

			if (l == 6)//A la fin de l'acquisition de donnée quand on a recu le dernier #
			{
				bool receptionok = false;
				if (somme1.ToString() == valCS1)
				{
					textBox9.AppendText(newLine + dataascii1);
					serialPort1.Write("v");//donnees recu et verifiees
					numerotramevrai = 1;//permettra par la suite de selectioner la trame 1 et donc les données ecrites dans la database1 comme une trame juste qui peut être ecrites dans le fichier Excel
					 receptionok =true;
					
				}
				else if (somme2.ToString() == valCS2)
				{
					textBox9.AppendText(newLine + dataascii2);
					serialPort1.Write("v");
					numerotramevrai = 2;
					receptionok = true;
				}
				else if (somme3.ToString() == valCS3){
					textBox9.AppendText(newLine + dataascii3);
					serialPort1.Write("v");
					numerotramevrai = 3;
					receptionok = true;
				}
				
				if ((receptionok == true)&&(modeauto == true)) {//en mode auto seulement on ecrit toutes les data dans excel
					Thread t = new Thread(new ThreadStart(EcrituredataThread));//on va ecrire la ligne de donnée dans le excel ou gridview
					t.IsBackground = true;
					t.Start();
					joursansreception = 0;//une fois le thread d'ecriture des données dans excel réalisé on constate que l'acquisition c'est bien passer aujourdhui
					serialPort1.Write("z");//On peut alors fermer la passerelle 
				}
				k = 0;
			}
				////////////   Le code qui suit concerne la partie debug de l'application :

				verificationchecksomme =false;//C'est 2 bool vont permettre d'indiquer quand on aura 
				verificationnbrligne=false;

				textBox7.Clear();//Dans la textbox7 on ecrit tous les checksomme et nombre de ligne recu et filtrer 
				textBox7.AppendText(newLine + "Calculer :"+ newLine + "  Checksome :"+ newLine + "1 : " + somme1 + newLine + "2 : " + somme2 + newLine + "3 : " + somme3);
				textBox7.AppendText(newLine + "   Nbr de lignes :" + newLine + "1 : " + nbrligne1 + newLine + "2 : " + nbrligne2 + newLine + "3 : " + nbrligne3);
				textBox7.AppendText(newLine + "Recu/Filtrer :" + newLine + "  Checksome :" + newLine + "1 : " + valCS1 + newLine + "2 : " + valCS2 + newLine + "3 : " + valCS3);
				textBox7.AppendText(newLine + "   Nbr de lignes :" + newLine + "1 : " + valnbrL1 + newLine + "2 : " + valnbrL2 + newLine + "3 : " + valnbrL3);
				textBox6.Clear();
				///verification des calculs des donnes recus

				//Le code suivant permet de vérifier si les checksommes des trames sont identiques 
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
				if ((valCS1 != valCS2) || (valCS2 != valCS3) || (valCS1 != valCS3))
				{ textBox6.AppendText(newLine + " ! Les checksommes recues sont differents ");
				}

			if ((valCS1 == valCS2) || (valCS1 == valCS3)) { checksommerecuvrai = valCS1;
				if (verificationchecksomme == true)
				{
					if (checksommerecuvrai == checksommecalculervrai.ToString())
					{
						textBox6.AppendText(newLine + "Checksommes recus et verifier : " + checksommerecuvrai);
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
						label8.BackColor = Color.Green;
					}
				}
			}
					if ((valCS1 != valCS2) && (valCS2 != valCS3) && (valCS1 != valCS3)) { 
				textBox6.AppendText(newLine + "!!! Aucuns checksums recus identiques"); 
					}

				else { textBox6.AppendText(newLine + "Les 3 checksommes recues sont identiques !"+ valCS1);
					checksommerecuvrai = valCS1;
					if (verificationchecksomme == true) {
						if (checksommerecuvrai == checksommecalculervrai.ToString()) {
							textBox6.AppendText(newLine + "Checksommes recus et verifier : " + checksommerecuvrai);
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
								textBox6.AppendText(newLine + "Nbr de lignes recus et verifier : " + nbrlignecalculervrai);
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

		}//end serialport1_DataReceived
		private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
		{
			if (serialPort1.PortName != comboBox1.SelectedItem.ToString())//si on choisi de changer de COM
			{
				// Fermeture du port COM actuel :
				if (serialPort1.IsOpen == true)
					serialPort1.Close();
				// Prise en compte du port COM choisi par l'utilisateur :
				serialPort1.PortName = comboBox1.SelectedItem.ToString();
				
				// Ouverture du nouveau port COM
				if (serialPort1.IsOpen == false)
					try { serialPort1.Open(); }
					catch ( Exception e3)
                    {
						MessageBox.Show(e3.Message, "Erreur !");
                    }
			}
		}
		private void button3_Click(object sender, EventArgs e)//Bouton test
		{
			serialPort1.Write("t");
			testHorloge = true;
		}
		private void button4_Click(object sender, EventArgs e)//Bouton data
		{
			j = 0;//on reintialise le cpt qui compte le nbr de trame de donnes <"..."> envoyer
			serialPort1.Write("l");	
		}
		private void button5_Click(object sender, EventArgs e)//Bouton compteur
		{
			serialPort1.Write("c");
		}
		private void button9_Click(object sender, EventArgs e)//Bouton GPS
		{
			serialPort1.Write("g");
		}
        private void button6_Click(object sender, EventArgs e)//Bouton clear associer à la textbox1
        {
			textBox1.Clear();
        }
        private void button7_Click(object sender, EventArgs e)
        {
			textBox2.Clear();
        }

		private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)//si on change le baud rate
        {
			
			serialPort1.DiscardInBuffer();//ou supprime les donnes encore dans le buffer pour eviter d'avoir des donnes fausse
			int newbaudrate = Convert.ToInt32(comboBox2.SelectedItem);
			if (newbaudrate > 0)
			{
				serialPort1.BaudRate = newbaudrate; 
				textBox8.Text=(serialPort1.BaudRate).ToString();
			}
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
			if (compteurc.Length == 4)//Il faut donc envoyer 4 caractéres aprés c
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
		
		private void EcrituredataThread()
        {
			Invoke(new ecrituredataDelegate(ecrituredata));//On doit invoquer la methode sinon il n'est pas possible de continuer a faire fonctionner le programme tout en ecrivant les donnée dans excel
        }
		private void ecrituredata()
		{
			if ((numerotramevrai == 1) || (numerotramevrai == 2) || (numerotramevrai == 3))//Si au moins 1 des trames a été vérifié
			{
				if (modeauto == true)
				{
					String pathfile = @datatxtpathfile;
					//par exemple String pathfile = @"C:\Users\grisoni\Desktop\";//ou  @"C:\Users\ROMAIN.000\Desktop\"
					String filename = "data_FLUNTR.txt";
					System.IO.File.AppendAllText(pathfile + filename, textBox9.Text);//on sauvegarde les données filtrer
				}
			}

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
						worksheet.Cells[m+ligneexcel, n] = STdataFLUNTR;//+ Ajouter la valeur de la derniére ligne
					}
					n = n - 9;//Sinon a chaque ligne on aurait ecrit 9 colones plus loin
				}
				ligneexcel += m;//on garde en mémoire la derniére ligne ecrite dans excel

				Properties.Settings.Default.nombredeligneexcel = ligneexcel;//on la sauvegarde dans les paramétre de l'application au cas ou celle-ci redemarrerer
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
						worksheet.Cells[m+ligneexcel, n] = STdataFLUNTR;//+  Ajouter la valeur de la derniére ligne
					}
					n = n - 9;
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
						worksheet.Cells[m+ligneexcel, n] = STdataFLUNTR;//+ Ajouter la valeur de la derniére ligne
					}
					n = n - 9;	
				}
				ligneexcel += m;//on enregistre le numero de la derniére ligne ecrite
				Properties.Settings.Default.nombredeligneexcel = ligneexcel;
				Properties.Settings.Default.Save();
			}
		}

        private void textBox13_KeyPress(object sender, KeyPressEventArgs e)//Permet d'envoyer 1 caractére personaliser à l'interface lointaine
        {
			string cptcommande = textBox13.Text;
			if (cptcommande.Length == 1)
			{
				serialPort1.Write(cptcommande);
				textBox5.Text = "Vous avez entrer : " + cptcommande;
				textBox13.Clear();
			}
		}

        private void button13_Click(object sender, EventArgs e)//Fermetture passerelle
        {
			serialPort1.Write("z");
		}

        private void button14_Click(object sender, EventArgs e)//ecriture manuelle dans un ficher texte
        {
			try
			{
				String pathfile = @datatxtpathfile;
				//Ce chemin d'accés doit être de la forme :@"C:\Users\grisoni\Desktop\";// @"C:\Users\ROMAIN.000\Desktop\"
				String filename = "data_FLUNTR.txt";
				System.IO.File.AppendAllText(pathfile + filename, textBox9.Text);//on sauvegarde les données filtrer//2>>9
				MessageBox.Show("Data has been saved to" + pathfile, "savefile");
				

			}
			catch (Exception e2)
			{
				MessageBox.Show(e2.Message, "error");
			}
		}

        private void button16_Click(object sender, EventArgs e)//choix de l'empacement du fichier
        {
			@datatxtpathfile = textBox14.Text;

			Properties.Settings.Default.EmplacementFichier = @datatxtpathfile;//Qu'on enregistre également dans les paramétres de l'application pour ne pas à avoir à le repréciser à chaque redemmarage
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

        private void button1_Click(object sender, EventArgs e)//Afficher fenetres pour debug les checksums et nbr lignes
        {
			panel4.Hide();//Utilisation d'un panel pour "cacher" les fenetres de debug car il n'est pas possible d'utiliser les 
			button1.Hide();//methodes .show ou .hide pour cacher/affciher directement les textbox 
		}

        private void button18_Click(object sender, EventArgs e)//fermer fenetres pour debug les checksums et nbr lignes
		{
			panel4.Show();
			button1.Show();
		}

        private void button19_Click(object sender, EventArgs e)//On affiche les fenetre de debug
		{
			panel3.Hide();
			button19.Hide();
		}

        private void button20_Click(object sender, EventArgs e)
        {

			panel3.Show();
			button19.Show();
		}

        private void button21_Click(object sender, EventArgs e)
        {
			System.Diagnostics.Process.Start(Application.ExecutablePath);
			Application.Exit();
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
