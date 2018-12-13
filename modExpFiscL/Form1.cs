using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Odbc;
using System.IO;

namespace modExpFiscL
{
    public partial class FrmExport : Form
    {
        // Initialisation des variables
        String portServeur = null;
        String adrServeur = null;
        String requetes = null;
        String paramConnexion = null;
        String loginPG = "postgres";
        String passPG = "pgpass";
        Boolean mesTests = true;

        OdbcConnection chaineConnexionListeBases = new OdbcConnection();
        OdbcCommand mesRequetes = new OdbcCommand();
        OdbcDataReader retourRequetes;

        public FrmExport()
        {
            InitializeComponent();
        }

        private void FrmExport_Load(object sender, EventArgs e)
        {
            // Remplissage des textboxs
            if (textBoxPort.Text == null)
            {
                textBoxPort.Text = "5432";
                portServeur = "5432";
            }
            else
            {
                portServeur = textBoxPort.Text;
            }

            if (textBoxServeur.Text == null)
            {
                textBoxServeur.Text = "127.0.0.1";
                adrServeur = "127.0.0.1";
            }
            else
            {
                adrServeur = textBoxServeur.Text;
            }

            // Initialisations des paramètres pour le passage de la requête
            requetes = "SELECT datname FROM pg_database WHERE datistemplate=FALSE and datname not like 'postgres'";
            paramConnexion = "Driver={PostgreSQL Unicode};Server=" + adrServeur + ";Port=" + portServeur + ";Database=postgres;Uid=" + loginPG + ";Pwd=" + passPG + ";";

            using (chaineConnexionListeBases = new OdbcConnection(paramConnexion))
            {
                try
                {
                    // Passage de la commande
                    mesRequetes = new OdbcCommand(requetes, chaineConnexionListeBases);
                    // Ouverture
                    chaineConnexionListeBases.Open();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erreur à la connexion au serveur pour la récupération de la liste des bases de données.");
                    mesTests = false;
                }
                finally
                {
                    // Exécution si pas d'exception
                    if (mesTests)
                    {
                        // Exeécution de la requête
                        retourRequetes = mesRequetes.ExecuteReader();

                        // Remplissage du comboBox
                        while (retourRequetes.Read())
                        {
                            comboBoxBase.Items.Add(retourRequetes["datname"]);
                        }

                        comboBoxBase.Text = comboBoxBase.Items[0].ToString();
                    }
                }
            }
        }

        private void textBoxServeur_KeyUp(object sender, KeyEventArgs e)
        {
            
            if ((e.KeyCode == Keys.Return) && (textBoxPort.Text != null))
            {
                // Initialsation des variables des valeurs en textBoxs
                portServeur = textBoxPort.Text;
                adrServeur = textBoxServeur.Text;

                // Initialisations des paramètres pour le passage de la requête
                requetes = "SELECT datname FROM pg_database WHERE datistemplate=FALSE and datname not like 'postgres'";
                paramConnexion = "Driver={PostgreSQL Unicode};Server=" + adrServeur + ";Port=" + portServeur + ";Database=postgres;Uid=" + loginPG + ";Pwd=" + passPG + ";";

                // Si le port et l'adresse sont bien renseignés
                // Mise à jour du comboBox
                // Connexion à la base
                if ((portServeur != null) && (adrServeur != null))
                {
                    using (chaineConnexionListeBases = new OdbcConnection(paramConnexion))
                    {
                        try
                        {
                            // Passage de la commande
                            mesRequetes = new OdbcCommand(requetes, chaineConnexionListeBases);
                            // Ouverture
                            chaineConnexionListeBases.Open();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Erreur à la connexion à la base suite à la modification de l'adresse ou du nom du serveur, vérifiez le paramètre.");
                        }
                        finally
                        {
                            // Exécution si pas d'exception
                            if (mesTests)
                            {
                                // Exeécution de la requête
                                retourRequetes = mesRequetes.ExecuteReader();

                                // On vide la comboBox
                                comboBoxBase.Items.Clear();

                                // Remplissage du comboBox
                                while (retourRequetes.Read())
                                {
                                    comboBoxBase.Items.Add(retourRequetes["datname"]);
                                }

                                // Fermeture de la connexion
                                retourRequetes.Close();

                                // On récupère la première valeur
                                comboBoxBase.Text = comboBoxBase.Items[0].ToString();
                            }
                        }
                    }
                }
            }
        }

        private void textBoxPort_KeyUp(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Return) && (textBoxServeur.Text != null))
            {
                // Initialisation des variables
                mesTests = true;

                // Récupération des valeurs en textBoxs
                portServeur = textBoxPort.Text;
                adrServeur = textBoxServeur.Text;

                // Initialisations des paramètres pour le passage de la requête
                requetes = "SELECT datname FROM pg_database WHERE datistemplate=FALSE and datname not like 'postgres'";
                paramConnexion = "Driver={PostgreSQL UNICODE};Server=" + adrServeur + ";Port=" + portServeur + ";Database=postgres;Uid=" + loginPG + ";Pwd=" + passPG + ";";

                // Si le port et l'adresse sont bien renseignés
                // Mise à jour du comboBox
                // Connexion à la base
                if ((portServeur != null) && (adrServeur != null))
                {
                    using (chaineConnexionListeBases = new OdbcConnection(paramConnexion))
                    {
                        try
                        {
                            // Passage de la commande
                            mesRequetes = new OdbcCommand(requetes, chaineConnexionListeBases);
                            // Ouverture
                            chaineConnexionListeBases.Open();
                        }
                        catch (Exception ex)
                        {
                            mesTests = false;
                            MessageBox.Show("Erreur à la connexion à la base suite à la modification du port, vérifiez le paramètre.");
                        }
                        finally
                        {
                            // Exécution si pas d'exception
                            if (mesTests)
                            {
                                // Exeécution de la requête
                                retourRequetes = mesRequetes.ExecuteReader();

                                // On vide la comboBox
                                comboBoxBase.Items.Clear();

                                // Remplissage du comboBox
                                while (retourRequetes.Read())
                                {
                                    comboBoxBase.Items.Add(retourRequetes["datname"]);
                                }

                                // Fermeture de la connexion
                                retourRequetes.Close();

                                // On récupère la première valeur
                                comboBoxBase.Text = comboBoxBase.Items[0].ToString();
                            }
                        }
                    }
                }
            }
        }

        private void btnExtraire_Click(object sender, EventArgs e)
        {
            // Déclaration des variables
            String requete1 = null;
            String requeteTVA1 = null;
            String requeteTVA2 = null;
            String requeteTVA3 = null;
            String requeteTVA4 = null;
            String requeteART = null;
            StreamWriter sw = new StreamWriter(Application.StartupPath + "\\extraction_prodevis.csv");
            OdbcCommand maRequete1 = new OdbcCommand();
            OdbcCommand maRequeteTVA1 = new OdbcCommand();
            OdbcCommand maRequeteTVA2 = new OdbcCommand();
            OdbcCommand maRequeteTVA3 = new OdbcCommand();
            OdbcCommand maRequeteTVA4 = new OdbcCommand();
            OdbcCommand maRequeteART = new OdbcCommand();
            OdbcDataReader retourRequete1;
            OdbcDataReader retourRequeteTVA1;
            OdbcDataReader retourRequeteTVA2;
            OdbcDataReader retourRequeteTVA3;
            OdbcDataReader retourRequeteTVA4;
            OdbcDataReader retourRequeteART;
            int i = 0;

            paramConnexion = "Driver={PostgreSQL UNICODE};Server=" + textBoxServeur.Text + ";Port=" + textBoxPort.Text + ";Database=" + comboBoxBase.Text + ";Uid=" + loginPG + ";Pwd=" + passPG + ";";

            requete1 = "SELECT f.id_facture AS ID, f.numdoc AS NUMFACT, dtdoc AS DATFAC, CONCAT(lv.liste, ' ', cc.nom, ' ', cc.prenom) AS NOM, f.totalttc AS TOTTTC, f.taxe_taux1 AS TXTAXE1, taxe_taux2 AS TXTAXE2, taxe_taux3 AS TXTAXE3, taxe_taux4 AS TXTAXE4 FROM facture f, client_coordonnees cc, listvoca lv WHERE cc.titre_id = lv.id AND f.client_coordonnees_id = cc.id_client_coordonnees";
            
            if ((portServeur != null) && (adrServeur != null))
            {
                using (chaineConnexionListeBases = new OdbcConnection(paramConnexion))
                {
                    try
                    {
                        // Passage de la commande
                        maRequete1 = new OdbcCommand(requete1, chaineConnexionListeBases);
                        // Ouverture
                        chaineConnexionListeBases.Open();
                    }
                    catch (Exception ex)
                    {
                        mesTests = false;
                        MessageBox.Show("Erreur à la connexion à la base de données sélectionnée.");
                        Application.Exit();
                    }
                    finally
                    {
                        // Exécution si pas d'exception
                        if (mesTests)
                        {
                            // Exeécution de la requête
                            retourRequete1 = maRequete1.ExecuteReader();
                            try
                            {
                                // Remplissage du comboBox
                                while (retourRequete1.Read())
                                {
                                    sw.Write(retourRequetes["NUMFACT"]);
                                    sw.Write(";");
                                    sw.Write(retourRequetes["DATFAC"]);
                                    sw.Write(";");
                                    sw.Write(retourRequetes["NOM"]);
                                    sw.Write(";");
                                    //sw.Write(retourRequetes["TOTTTC"]);

                                    requeteTVA1 = "SELECT taxe_num, SUM(ROUND((prixventerpp*qte)::numeric,2))+SUM(ROUND((prixventepre2rpp*qte)::numeric,2)) AS PROD, SUM(ROUND((prixventepre1rpp*qte)::numeric,2)) AS PREST, SUM(ROUND((prix_ecotaxe*qte)::numeric,2)) AS ECO FROM ligne_affaire_facture WHERE produit_type = 'P' AND FACTURE_id=" + retourRequetes["ID"] + " AND taxe_num=1 GROUP BY taxe_num"
                                    requeteTVA2 = "SELECT taxe_num, SUM(ROUND((prixventerpp*qte)::numeric,2))+SUM(ROUND((prixventepre2rpp*qte)::numeric,2)) AS PROD, SUM(ROUND((prixventepre1rpp*qte)::numeric,2)) AS PREST, SUM(ROUND((prix_ecotaxe*qte)::numeric,2)) AS ECO FROM ligne_affaire_facture WHERE produit_type = 'P' AND FACTURE_id=" + retourRequetes["ID"] + " AND taxe_num=2 GROUP BY taxe_num"
                                    requeteTVA3 = "SELECT taxe_num, SUM(ROUND((prixventerpp*qte)::numeric,2))+SUM(ROUND((prixventepre2rpp*qte)::numeric,2)) AS PROD, SUM(ROUND((prixventepre1rpp*qte)::numeric,2)) AS PREST, SUM(ROUND((prix_ecotaxe*qte)::numeric,2)) AS ECO FROM ligne_affaire_facture WHERE produit_type = 'P' AND FACTURE_id=" + retourRequetes["ID"] + " AND taxe_num=3 GROUP BY taxe_num"
                                    requeteTVA4 = "SELECT taxe_num, SUM(ROUND((prixventerpp*qte)::numeric,2))+SUM(ROUND((prixventepre2rpp*qte)::numeric,2)) AS PROD, SUM(ROUND((prixventepre1rpp*qte)::numeric,2)) AS PREST, SUM(ROUND((prix_ecotaxe*qte)::numeric,2)) AS ECO FROM ligne_affaire_facture WHERE produit_type = 'P' AND FACTURE_id=" + retourRequetes["ID"] + " AND taxe_num=4 GROUP BY taxe_num"
                                    requeteART = "SELECT id_ligne_affaire_fact AS ID, designation AS DESIG FROM ligne_affaire_facture, CONCAT (options_produit, ' - ', options_prest) AS OPTIONS WHERE produit_type = 'P' AND FACTURE_id=" + retourRequetes["ID"];


                                    try
                                    {
                                        // Passage de la commande
                                        maRequeteTVA1 = new OdbcCommand(requete1, chaineConnexionListeBases);
                                        // Ouverture
                                        chaineConnexionListeBases.Open();
                                    }
                                    catch (Exception ex)
                                    {
                                        mesTests = false;
                                        MessageBox.Show("Erreur à la connexion à la base de données sélectionnée pour lister les produits.");
                                        Application.Exit();
                                    }
                                    finally
                                    {
                                        retourRequeteART = maRequeteTVA1.ExecuteReader();

                                        i = 0;

                                        while (retourRequeteART.Read())
                                        {
                                            if (i == 0)
                                            {
                                                sw.Write("Article " + i + " : " + retourRequeteART["DESIG"]);
                                            }
                                            else
                                            {
                                                sw.Write(" - Article " + i + " : " + retourRequeteART["DESIG"]);
                                            }
                                            i++;
                                        }

                                        i = 0;

                                        while (retourRequeteART.Read())
                                        {
                                            if (i == 0)
                                            {
                                                sw.Write("Options article " + i + " : " + retourRequeteART["DESIGNATION"]);
                                            }
                                            else
                                            {
                                                sw.Write(" - Options article " + i + " : " + retourRequeteART["OPTIONS"]);
                                            }
                                            i++;
                                        }


                                        retourRequeteTVA1 = maRequeteTVA1.ExecuteReader();
                                    }
                                }
                            }
                            catch(Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                            // Fermeture de la connexion
                            retourRequete1.Close();
                        }
                    }
                }
            }
        }
    }
}