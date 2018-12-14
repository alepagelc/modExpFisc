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
            requetes = "SELECT datname FROM pg_database WHERE datistemplate=FALSE and datname not like 'postgres' ORDER BY datname";
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
                requetes = "SELECT datname FROM pg_database WHERE datistemplate=FALSE and datname not like 'postgres' ORDER BY datname";
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
                requetes = "SELECT datname FROM pg_database WHERE datistemplate=FALSE and datname not like 'postgres' ORDER BY datname";
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
            OdbcConnection chaineConnexionReq1 = new OdbcConnection();
            OdbcConnection chaineConnexionReqART = new OdbcConnection();
            OdbcConnection chaineConnexionReqTVA1 = new OdbcConnection();
            OdbcConnection chaineConnexionReqTVA2 = new OdbcConnection();
            OdbcConnection chaineConnexionReqTVA3 = new OdbcConnection();
            OdbcConnection chaineConnexionReqTVA4 = new OdbcConnection();
            OdbcCommand maRequete1 = new OdbcCommand();
            OdbcCommand maRequeteTVA1 = new OdbcCommand();
            OdbcCommand maRequeteTVA2 = new OdbcCommand();
            OdbcCommand maRequeteTVA3 = new OdbcCommand();
            OdbcCommand maRequeteTVA4 = new OdbcCommand();
            OdbcCommand maRequeteART = new OdbcCommand();
            OdbcDataReader retourRequete1 = null;
            OdbcDataReader retourRequeteTVA1 = null;
            OdbcDataReader retourRequeteTVA2 = null;
            OdbcDataReader retourRequeteTVA3 = null;
            OdbcDataReader retourRequeteTVA4 = null;
            OdbcDataReader retourRequeteART = null;
            int i = 0;
            //Déclaration et initialisation des variables recueillant les valeurs issues des requêtes, taux de TVA encore inconnus
            double htProdTva1 = 0;
            double htProdTva2 = 0;
            double htProdTva3 = 0;
            double htProdTva4 = 0;
            double htPrestTva1 = 0;
            double htPrestTva2 = 0;
            double htPrestTva3 = 0;
            double htPrestTva4 = 0;
            //Déclaration et initialisation des variables recueillant les valeurs issues des tests, TVA connues après tests
            double totTva20 = 0;
            double totTva10 = 0;
            double totTva55 = 0;
            double totHtProd20 = 0;
            double totHtProd10 = 0;
            double totHtProd55 = 0;
            double totHtProd0 = 0;
            double totHtPrest20 = 0;
            double totHtPrest10 = 0;
            double totHtPrest55 = 0;
            double totHtPrest0 = 0;

            paramConnexion = "Driver={PostgreSQL UNICODE};Server=" + textBoxServeur.Text + ";Port=" + textBoxPort.Text + ";Database=" + comboBoxBase.Text + ";Uid=" + loginPG + ";Pwd=" + passPG + ";";

            requete1 = "SELECT f.id_facture AS ID, f.numdoc AS NUMFACT, dtdoc AS DATFAC, lv.liste || ' ' || cc.nom || ' ' || cc.prenom AS NOM, f.totalttc AS TOTTTC, f.taxe_taux1 AS TXTAXE1, taxe_taux2 AS TXTAXE2, taxe_taux3 AS TXTAXE3, taxe_taux4 AS TXTAXE4 FROM facture f, client_coordonnees cc, listvoca lv WHERE cc.titre_id = lv.id AND f.client_coordonnees_id = cc.id_client_coordonnees ORDER BY DATFAC";

            if ((portServeur != null) && (adrServeur != null))
            {
                using (chaineConnexionReq1 = new OdbcConnection(paramConnexion))
                {
                    try
                    {
                        // Passage de la commande
                        maRequete1 = new OdbcCommand(requete1, chaineConnexionReq1);
                        // Ouverture
                        chaineConnexionReq1.Open();
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
                            sw.Write("Numéro de la facture");
                            sw.Write(";");
                            sw.Write("Date de la facture");
                            sw.Write(";");
                            sw.Write("Nom du client");
                            sw.Write(";");
                            sw.Write("Listes articles");
                            sw.Write(";");
                            sw.Write("Listes options");
                            sw.Write(";");
                            sw.Write("Total TVA 20%");
                            sw.Write(";");
                            sw.Write("Total TVA 10%");
                            sw.Write(";");
                            sw.Write("Total TVA 5,5%");
                            sw.Write(";");
                            sw.Write("Total Produit 20%");
                            sw.Write(";");
                            sw.Write("Total Produit 10%");
                            sw.Write(";");
                            sw.Write("Total Produit 5,5%");
                            sw.Write(";");
                            sw.Write("Total Produit 0%");
                            sw.Write(";");
                            sw.Write("Total Prestation 20%");
                            sw.Write(";");
                            sw.Write("Total Prestation 10%");
                            sw.Write(";");
                            sw.Write("Total Prestation 5,5%");
                            sw.Write(";");
                            sw.Write("Total Prestation 0%");
                            sw.Write(";");
                            sw.Write("Total TTC");

                            // Exeécution de la requête
                            retourRequete1 = maRequete1.ExecuteReader();
                            try
                            {
                                // Remplissage du comboBox
                                while (retourRequete1.Read())
                                {
                                    sw.Write(retourRequete1["NUMFACT"]);
                                    sw.Write(";");
                                    sw.Write(retourRequete1["DATFAC"]);
                                    sw.Write(";");
                                    sw.Write(retourRequete1["NOM"]);
                                    sw.Write(";");

                                    requeteTVA1 = "SELECT taxe_num, SUM(ROUND((prixventerpp*qte)::numeric,2))+SUM(ROUND((prixventepre2rpp*qte)::numeric,2)) AS PROD, SUM(ROUND((prixventepre1rpp*qte)::numeric,2)) AS PREST, SUM(ROUND((prix_ecotaxe*qte)::numeric,2)) AS ECO FROM ligne_affaire_facture WHERE produit_type <> 'S' AND produit_type <> 'RC' AND FACTURE_id=" + retourRequete1["ID"] + " AND taxe_num=1 GROUP BY taxe_num";
                                    requeteTVA2 = "SELECT taxe_num, SUM(ROUND((prixventerpp*qte)::numeric,2))+SUM(ROUND((prixventepre2rpp*qte)::numeric,2)) AS PROD, SUM(ROUND((prixventepre1rpp*qte)::numeric,2)) AS PREST, SUM(ROUND((prix_ecotaxe*qte)::numeric,2)) AS ECO FROM ligne_affaire_facture WHERE produit_type <> 'S' AND produit_type <> 'RC' AND FACTURE_id=" + retourRequete1["ID"] + " AND taxe_num=2 GROUP BY taxe_num";
                                    requeteTVA3 = "SELECT taxe_num, SUM(ROUND((prixventerpp*qte)::numeric,2))+SUM(ROUND((prixventepre2rpp*qte)::numeric,2)) AS PROD, SUM(ROUND((prixventepre1rpp*qte)::numeric,2)) AS PREST, SUM(ROUND((prix_ecotaxe*qte)::numeric,2)) AS ECO FROM ligne_affaire_facture WHERE produit_type <> 'S' AND produit_type <> 'RC' AND FACTURE_id=" + retourRequete1["ID"] + " AND taxe_num=3 GROUP BY taxe_num";
                                    requeteTVA4 = "SELECT taxe_num, SUM(ROUND((prixventerpp*qte)::numeric,2))+SUM(ROUND((prixventepre2rpp*qte)::numeric,2)) AS PROD, SUM(ROUND((prixventepre1rpp*qte)::numeric,2)) AS PREST, SUM(ROUND((prix_ecotaxe*qte)::numeric,2)) AS ECO FROM ligne_affaire_facture WHERE produit_type <> 'S' AND produit_type <> 'RC' AND FACTURE_id=" + retourRequete1["ID"] + " AND taxe_num=4 GROUP BY taxe_num";
                                    requeteART = "SELECT id_ligne_affaire_facture AS ID, options_produit || ' - ' || options_prest AS OPTIONS, designation AS DESIG FROM ligne_affaire_facture WHERE produit_type <> 'S' AND produit_type <> 'RC' AND FACTURE_id=" + retourRequete1["ID"];


                                    // CONNEXION POUR LES ARTICLES ET OPTIONS
                                    using (chaineConnexionReqART = new OdbcConnection(paramConnexion))
                                    {
                                        try
                                        {
                                            // Passage de la commande
                                            maRequeteART = new OdbcCommand(requeteART, chaineConnexionReqART);
                                            // Ouverture
                                            chaineConnexionReqART.Open();
                                        }
                                        catch (Exception ex)
                                        {
                                            mesTests = false;
                                            MessageBox.Show("Erreur à la connexion à la base de données sélectionnée pour lister les produits.");
                                            Application.Exit();
                                        }
                                        finally
                                        {
                                            try
                                            {
                                                retourRequeteART = maRequeteART.ExecuteReader();
                                            }
                                            catch (Exception ex)
                                            {
                                                MessageBox.Show("Erreur à l'éxécution de la requête pour lister les produits : " + ex);
                                                Application.Exit();
                                            }
                                            finally
                                            {

                                                try
                                                {
                                                    i = 1;

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
                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show("Erreur à la récupération des articles");
                                                    Application.Exit();
                                                }
                                                finally
                                                {
                                                    sw.Write(";");
                                                }

                                                try
                                                {
                                                    i = 1;

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
                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show("Erreur à la récupération des options");
                                                    Application.Exit();
                                                }
                                                finally
                                                {
                                                    sw.Write(";");
                                                }
                                            }

                                            using (chaineConnexionReqTVA1 = new OdbcConnection(paramConnexion))
                                            {
                                                try
                                                {
                                                    // Passage de la commande
                                                    maRequeteTVA1 = new OdbcCommand(requeteTVA1, chaineConnexionListeBases);
                                                    // Ouverture
                                                    chaineConnexionReqTVA1.Open();
                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show(ex.ToString());
                                                }
                                                finally
                                                {
                                                    try
                                                    {
                                                        retourRequeteTVA1 = maRequeteTVA1.ExecuteReader();
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        MessageBox.Show(ex.ToString());
                                                        Application.Exit();
                                                    }
                                                    finally
                                                    {
                                                        while(retourRequeteTVA1.Read())
                                                        {
                                                            htProdTva1 = Convert.ToDouble(retourRequeteTVA1["PROD"]);
                                                            htPrestTva1 = Convert.ToDouble(retourRequeteTVA1["PREST"]);
                                                        }
                                                    }
                                                }
                                            }

                                            using (chaineConnexionReqTVA2 = new OdbcConnection(paramConnexion))
                                            {
                                                try
                                                {
                                                    // Passage de la commande
                                                    maRequeteTVA1 = new OdbcCommand(requeteTVA2, chaineConnexionListeBases);
                                                    // Ouverture
                                                    chaineConnexionReqTVA2.Open();
                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show(ex.ToString());
                                                    Application.Exit();
                                                }
                                                finally
                                                {
                                                    try
                                                    {
                                                        retourRequeteTVA2 = maRequeteTVA2.ExecuteReader();
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        MessageBox.Show(ex.ToString());
                                                        Application.Exit();
                                                    }
                                                    finally
                                                    {
                                                        while (retourRequeteTVA2.Read())
                                                        {
                                                            htProdTva2 = Convert.ToDouble(retourRequeteTVA2["PROD"]);
                                                            htPrestTva2 = Convert.ToDouble(retourRequeteTVA2["PREST"]);
                                                        }
                                                    }
                                                }

                                            }

                                            using (chaineConnexionReqTVA3 = new OdbcConnection(paramConnexion))
                                            {
                                                try
                                                {
                                                    // Passage de la commande
                                                    maRequeteTVA3 = new OdbcCommand(requeteTVA3, chaineConnexionListeBases);
                                                    // Ouverture
                                                    chaineConnexionReqTVA3.Open();
                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show(ex.ToString());
                                                    Application.Exit();
                                                }
                                                finally
                                                {
                                                    try
                                                    {
                                                        retourRequeteTVA3 = maRequeteTVA3.ExecuteReader();
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        MessageBox.Show(ex.ToString());
                                                        Application.Exit();
                                                    }
                                                    finally
                                                    {
                                                        while (retourRequeteTVA3.Read())
                                                        {
                                                            htProdTva3 = Convert.ToDouble(retourRequeteTVA3["PROD"]);
                                                            htPrestTva3 = Convert.ToDouble(retourRequeteTVA3["PREST"]);
                                                        }                                                       
                                                    }
                                                }
                                            }

                                            using (chaineConnexionReqTVA4 = new OdbcConnection(paramConnexion))
                                            {
                                                try
                                                {
                                                    // Passage de la commande
                                                    maRequeteTVA4 = new OdbcCommand(requeteTVA4, chaineConnexionListeBases);
                                                    // Ouverture
                                                    chaineConnexionReqTVA4.Open();
                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show(ex.ToString());
                                                    Application.Exit();
                                                }
                                                finally
                                                {
                                                    try
                                                    {
                                                        retourRequeteTVA4 = maRequeteTVA4.ExecuteReader();
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        MessageBox.Show(ex.ToString());
                                                        Application.Exit();
                                                    }
                                                    finally
                                                    {
                                                        while(retourRequeteTVA4.Read())
                                                        {
                                                            htProdTva4 = Convert.ToDouble(retourRequeteTVA4["PROD"]);
                                                            htPrestTva4 = Convert.ToDouble(retourRequeteTVA4["PREST"]);
                                                        }                                                        
                                                    }
                                                }
                                            }

                                            retourRequeteTVA1.Close();
                                            retourRequeteTVA2.Close();
                                            retourRequeteTVA3.Close();
                                            retourRequeteTVA4.Close();
                                            retourRequeteART.Close();
                                        }
                                    }

                                    //Traitements ventes taux 1
                                    switch (Convert.ToDouble(retourRequete1["TXTAXE1"]))
                                    {
                                        case 20:
                                            totHtProd20 = htProdTva1;
                                            totHtPrest20 = htPrestTva1;
                                            totTva20 = (htProdTva1 + htPrestTva1) + 0.2;
                                            break;
                                        case 10:
                                            totHtProd10 = htProdTva1;
                                            totHtPrest10 = htPrestTva1;
                                            totTva10 = (htProdTva1 + htPrestTva1) + 0.1;
                                            break;
                                        case 5.5:
                                            totHtProd55 = htProdTva1;
                                            totHtPrest55 = htPrestTva1;
                                            totTva55 = (htProdTva1 + htPrestTva1) + 0.055;
                                            break;
                                        case 0:
                                            totHtProd0 = htProdTva1;
                                            totHtPrest0 = htPrestTva1;
                                            break;
                                        default:
                                            break;
                                    }

                                    //Traitements ventes taux 2
                                    switch (Convert.ToDouble(retourRequete1["TXTAXE2"]))
                                    {
                                        case 20:
                                            totHtProd20 = htProdTva2;
                                            totHtPrest20 = htPrestTva2;
                                            totTva20 = (htProdTva2 + htPrestTva2) + 0.2;
                                            break;
                                        case 10:
                                            totHtProd10 = htProdTva2;
                                            totHtPrest10 = htPrestTva2;
                                            totTva10 = (htProdTva2 + htPrestTva2) + 0.1;
                                            break;
                                        case 5.5:
                                            totHtProd55 = htProdTva2;
                                            totHtPrest55 = htPrestTva2;
                                            totTva55 = (htProdTva2 + htPrestTva2) + 0.055;
                                            break;
                                        case 0:
                                            totHtProd0 = htProdTva2;
                                            totHtPrest0 = htPrestTva2;
                                            break;
                                        default:
                                            break;
                                    }

                                    //Traitements ventes taux 3
                                    switch (Convert.ToDouble(retourRequete1["TXTAXE3"]))
                                    {
                                        case 20:
                                            totHtProd20 = htProdTva3;
                                            totHtPrest20 = htPrestTva3;
                                            totTva20 = (htProdTva3 + htPrestTva3) + 0.2;
                                            break;
                                        case 10:
                                            totHtProd10 = htProdTva3;
                                            totHtPrest10 = htPrestTva3;
                                            totTva10 = (htProdTva3 + htPrestTva3) + 0.1;
                                            break;
                                        case 5.5:
                                            totHtProd55 = htProdTva3;
                                            totHtPrest55 = htPrestTva3;
                                            totTva55 = (htProdTva3 + htPrestTva3) + 0.055;
                                            break;
                                        case 0:
                                            totHtProd0 = htProdTva3;
                                            totHtPrest0 = htPrestTva3;
                                            break;
                                        default:
                                            break;
                                    }

                                    //Traitements ventes taux 4
                                    switch (Convert.ToDouble(retourRequete1["TXTAXE4"]))
                                    {
                                        case 20:
                                            totHtProd20 = htProdTva4;
                                            totHtPrest20 = htPrestTva4;
                                            totTva20 = (htProdTva4 + htPrestTva4) + 0.2;
                                            break;
                                        case 10:
                                            totHtProd10 = htProdTva4;
                                            totHtPrest10 = htPrestTva4;
                                            totTva10 = (htProdTva4 + htPrestTva4) + 0.1;
                                            break;
                                        case 5.5:
                                            totHtProd55 = htProdTva4;
                                            totHtPrest55 = htPrestTva4;
                                            totTva55 = (htProdTva4 + htPrestTva4) + 0.055;
                                            break;
                                        case 0:
                                            totHtProd0 = htProdTva4;
                                            totHtPrest0 = htPrestTva4;
                                            break;
                                        default:
                                            break;
                                    }

                                    sw.Write(Convert.ToString(totTva20));
                                    sw.Write(";");
                                    sw.Write(Convert.ToString(totTva10));
                                    sw.Write(";");
                                    sw.Write(Convert.ToString(totTva55));
                                    sw.Write(";");
                                    sw.Write(Convert.ToString(totHtProd20));
                                    sw.Write(";");
                                    sw.Write(Convert.ToString(totHtProd10));
                                    sw.Write(";");
                                    sw.Write(Convert.ToString(totHtProd55));
                                    sw.Write(";");
                                    sw.Write(Convert.ToString(totHtProd0));
                                    sw.Write(";");
                                    sw.Write(Convert.ToString(totHtPrest20));
                                    sw.Write(";");
                                    sw.Write(Convert.ToString(totHtPrest10));
                                    sw.Write(";");
                                    sw.Write(Convert.ToString(totHtPrest55));
                                    sw.Write(";");
                                    sw.Write(Convert.ToString(totHtPrest0));
                                    sw.Write(";");
                                    sw.Write(retourRequete1["TOTTTC"]);                                       
                                }
                                sw.Write(Environment.NewLine);
                            }
                            catch(Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }
                            finally
                            {
                                // Fermeture de la connexion
                                retourRequete1.Close();
                            }
                        }
                    }
                }
            }

            requete1 = "SELECT f.id_facture_arch AS ID, f.numdoc AS NUMFACT, dtdoc AS DATFAC, lv.liste || ' ' || cc.nom || ' ' || cc.prenom AS NOM, f.totalttc AS TOTTTC, f.taxe_taux1 AS TXTAXE1, taxe_taux2 AS TXTAXE2, taxe_taux3 AS TXTAXE3, taxe_taux4 AS TXTAXE4 FROM facture_arch f, client_coordonnees cc, listvoca lv WHERE cc.titre_id = lv.id AND f.client_coordonnees_id = cc.id_client_coordonnees ORDER BY DATFAC";

            if ((portServeur != null) && (adrServeur != null))
            {
                using (chaineConnexionReq1 = new OdbcConnection(paramConnexion))
                {
                    try
                    {
                        // Passage de la commande
                        maRequete1 = new OdbcCommand(requete1, chaineConnexionReq1);
                        // Ouverture
                        chaineConnexionReq1.Open();
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
                                    sw.Write(retourRequete1["NUMFACT"]);
                                    sw.Write(";");
                                    sw.Write(retourRequete1["DATFAC"]);
                                    sw.Write(";");
                                    sw.Write(retourRequete1["NOM"]);
                                    sw.Write(";");

                                    requeteTVA1 = "SELECT taxe_num, SUM(ROUND((prixventerpp*qte)::numeric,2))+SUM(ROUND((prixventepre2rpp*qte)::numeric,2)) AS PROD, SUM(ROUND((prixventepre1rpp*qte)::numeric,2)) AS PREST, SUM(ROUND((prix_ecotaxe*qte)::numeric,2)) AS ECO FROM ligne_affaire_facture_arch WHERE produit_type <> 'S' AND produit_type <> 'RC' AND facture_arch_id=" + retourRequete1["ID"] + " AND taxe_num=1 GROUP BY taxe_num";
                                    requeteTVA2 = "SELECT taxe_num, SUM(ROUND((prixventerpp*qte)::numeric,2))+SUM(ROUND((prixventepre2rpp*qte)::numeric,2)) AS PROD, SUM(ROUND((prixventepre1rpp*qte)::numeric,2)) AS PREST, SUM(ROUND((prix_ecotaxe*qte)::numeric,2)) AS ECO FROM ligne_affaire_facture_arch WHERE produit_type <> 'S' AND produit_type <> 'RC' AND facture_arch_id=" + retourRequete1["ID"] + " AND taxe_num=2 GROUP BY taxe_num";
                                    requeteTVA3 = "SELECT taxe_num, SUM(ROUND((prixventerpp*qte)::numeric,2))+SUM(ROUND((prixventepre2rpp*qte)::numeric,2)) AS PROD, SUM(ROUND((prixventepre1rpp*qte)::numeric,2)) AS PREST, SUM(ROUND((prix_ecotaxe*qte)::numeric,2)) AS ECO FROM ligne_affaire_facture_arch WHERE produit_type <> 'S' AND produit_type <> 'RC' AND facture_arch_id=" + retourRequete1["ID"] + " AND taxe_num=3 GROUP BY taxe_num";
                                    requeteTVA4 = "SELECT taxe_num, SUM(ROUND((prixventerpp*qte)::numeric,2))+SUM(ROUND((prixventepre2rpp*qte)::numeric,2)) AS PROD, SUM(ROUND((prixventepre1rpp*qte)::numeric,2)) AS PREST, SUM(ROUND((prix_ecotaxe*qte)::numeric,2)) AS ECO FROM ligne_affaire_facture_arch WHERE produit_type <> 'S' AND produit_type <> 'RC' AND facture_arch_id=" + retourRequete1["ID"] + " AND taxe_num=4 GROUP BY taxe_num";
                                    requeteART = "SELECT id_ligne_affaire_facture AS ID, options_produit || ' - ' || options_prest AS OPTIONS, designation AS DESIG FROM ligne_affaire_facture WHERE produit_type <> 'S' AND produit_type <> 'RC' AND facture_arch_id=" + retourRequete1["ID"];


                                    // CONNEXION POUR LES ARTICLES ET OPTIONS
                                    using (chaineConnexionReqART = new OdbcConnection(paramConnexion))
                                    {
                                        try
                                        {
                                            // Passage de la commande
                                            maRequeteART = new OdbcCommand(requeteART, chaineConnexionReqART);
                                            // Ouverture
                                            chaineConnexionReqART.Open();
                                        }
                                        catch (Exception ex)
                                        {
                                            mesTests = false;
                                            MessageBox.Show("Erreur à la connexion à la base de données sélectionnée pour lister les produits.");
                                            Application.Exit();
                                        }
                                        finally
                                        {
                                            try
                                            {
                                                retourRequeteART = maRequeteART.ExecuteReader();
                                            }
                                            catch (Exception ex)
                                            {
                                                MessageBox.Show("Erreur à l'éxécution de la requête pour lister les produits : " + ex);
                                                Application.Exit();
                                            }
                                            finally
                                            {

                                                try
                                                {
                                                    i = 1;

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
                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show("Erreur à la récupération des articles");
                                                    Application.Exit();
                                                }
                                                finally
                                                {
                                                    sw.Write(";");
                                                }

                                                try
                                                {
                                                    i = 1;

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
                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show("Erreur à la récupération des options");
                                                    Application.Exit();
                                                }
                                                finally
                                                {
                                                    sw.Write(";");
                                                }
                                            }

                                            using (chaineConnexionReqTVA1 = new OdbcConnection(paramConnexion))
                                            {
                                                try
                                                {
                                                    // Passage de la commande
                                                    maRequeteTVA1 = new OdbcCommand(requeteTVA1, chaineConnexionListeBases);
                                                    // Ouverture
                                                    chaineConnexionReqTVA1.Open();
                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show(ex.ToString());
                                                }
                                                finally
                                                {
                                                    try
                                                    {
                                                        retourRequeteTVA1 = maRequeteTVA1.ExecuteReader();
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        MessageBox.Show(ex.ToString());
                                                        Application.Exit();
                                                    }
                                                    finally
                                                    {
                                                        while(retourRequeteTVA1.Read())
                                                        {
                                                            htProdTva1 = Convert.ToDouble(retourRequeteTVA1["PROD"]);
                                                            htPrestTva1 = Convert.ToDouble(retourRequeteTVA1["PREST"]);
                                                        }                                                        
                                                    }
                                                }
                                            }

                                            using (chaineConnexionReqTVA2 = new OdbcConnection(paramConnexion))
                                            {
                                                try
                                                {
                                                    // Passage de la commande
                                                    maRequeteTVA1 = new OdbcCommand(requeteTVA2, chaineConnexionListeBases);
                                                    // Ouverture
                                                    chaineConnexionReqTVA2.Open();
                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show(ex.ToString());
                                                    Application.Exit();
                                                }
                                                finally
                                                {
                                                    try
                                                    {
                                                        retourRequeteTVA2 = maRequeteTVA2.ExecuteReader();
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        MessageBox.Show(ex.ToString());
                                                        Application.Exit();
                                                    }
                                                    finally
                                                    {
                                                        while(retourRequeteTVA2.Read())
                                                        {
                                                            htProdTva2 = Convert.ToDouble(retourRequeteTVA2["PROD"]);
                                                            htPrestTva2 = Convert.ToDouble(retourRequeteTVA2["PREST"]);
                                                        }                                                        
                                                    }
                                                }
                                            }

                                            using (chaineConnexionReqTVA3 = new OdbcConnection(paramConnexion))
                                            {
                                                try
                                                {
                                                    // Passage de la commande
                                                    maRequeteTVA3 = new OdbcCommand(requeteTVA3, chaineConnexionListeBases);
                                                    // Ouverture
                                                    chaineConnexionReqTVA3.Open();
                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show(ex.ToString());
                                                    Application.Exit();
                                                }
                                                finally
                                                {
                                                    try
                                                    {
                                                        retourRequeteTVA3 = maRequeteTVA3.ExecuteReader();
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        MessageBox.Show(ex.ToString());
                                                        Application.Exit();
                                                    }
                                                    finally
                                                    {
                                                        while(retourRequeteTVA3.Read())
                                                        {
                                                            htProdTva3 = Convert.ToDouble(retourRequeteTVA3["PROD"]);
                                                            htPrestTva3 = Convert.ToDouble(retourRequeteTVA3["PREST"]);
                                                        }                                                        
                                                    }
                                                }
                                            }

                                            using (chaineConnexionReqTVA4 = new OdbcConnection(paramConnexion))
                                            {
                                                try
                                                {
                                                    // Passage de la commande
                                                    maRequeteTVA4 = new OdbcCommand(requeteTVA4, chaineConnexionListeBases);
                                                    // Ouverture
                                                    chaineConnexionReqTVA4.Open();
                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show(ex.ToString());
                                                    Application.Exit();
                                                }
                                                finally
                                                {
                                                    try
                                                    {
                                                        retourRequeteTVA4 = maRequeteTVA4.ExecuteReader();
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        MessageBox.Show(ex.ToString());
                                                        Application.Exit();
                                                    }
                                                    finally
                                                    {
                                                        while(retourRequeteTVA4.Read())
                                                        {
                                                            htProdTva4 = Convert.ToDouble(retourRequeteTVA4["PROD"]);
                                                            htPrestTva4 = Convert.ToDouble(retourRequeteTVA4["PREST"]);
                                                        }                                                     
                                                    }
                                                }
                                            }

                                            retourRequeteTVA1.Close();
                                            retourRequeteTVA2.Close();
                                            retourRequeteTVA3.Close();
                                            retourRequeteTVA4.Close();
                                            retourRequeteART.Close();
                                        }
                                    }

                                    //Traitements ventes taux 1
                                    switch (Convert.ToDouble(retourRequete1["TXTAXE1"]))
                                    {
                                        case 20:
                                            totHtProd20 = htProdTva1;
                                            totHtPrest20 = htPrestTva1;
                                            totTva20 = (htProdTva1 + htPrestTva1) + 0.2;
                                            break;
                                        case 10:
                                            totHtProd10 = htProdTva1;
                                            totHtPrest10 = htPrestTva1;
                                            totTva10 = (htProdTva1 + htPrestTva1) + 0.1;
                                            break;
                                        case 5.5:
                                            totHtProd55 = htProdTva1;
                                            totHtPrest55 = htPrestTva1;
                                            totTva55 = (htProdTva1 + htPrestTva1) + 0.055;
                                            break;
                                        case 0:
                                            totHtProd0 = htProdTva1;
                                            totHtPrest0 = htPrestTva1;
                                            break;
                                        default:
                                            break;
                                    }

                                    //Traitements ventes taux 2
                                    switch (Convert.ToDouble(retourRequete1["TXTAXE2"]))
                                    {
                                        case 20:
                                            totHtProd20 = htProdTva2;
                                            totHtPrest20 = htPrestTva2;
                                            totTva20 = (htProdTva2 + htPrestTva2) + 0.2;
                                            break;
                                        case 10:
                                            totHtProd10 = htProdTva2;
                                            totHtPrest10 = htPrestTva2;
                                            totTva10 = (htProdTva2 + htPrestTva2) + 0.1;
                                            break;
                                        case 5.5:
                                            totHtProd55 = htProdTva2;
                                            totHtPrest55 = htPrestTva2;
                                            totTva55 = (htProdTva2 + htPrestTva2) + 0.055;
                                            break;
                                        case 0:
                                            totHtProd0 = htProdTva2;
                                            totHtPrest0 = htPrestTva2;
                                            break;
                                        default:
                                            break;
                                    }

                                    //Traitements ventes taux 3
                                    switch (Convert.ToDouble(retourRequete1["TXTAXE3"]))
                                    {
                                        case 20:
                                            totHtProd20 = htProdTva3;
                                            totHtPrest20 = htPrestTva3;
                                            totTva20 = (htProdTva3 + htPrestTva3) + 0.2;
                                            break;
                                        case 10:
                                            totHtProd10 = htProdTva3;
                                            totHtPrest10 = htPrestTva3;
                                            totTva10 = (htProdTva3 + htPrestTva3) + 0.1;
                                            break;
                                        case 5.5:
                                            totHtProd55 = htProdTva3;
                                            totHtPrest55 = htPrestTva3;
                                            totTva55 = (htProdTva3 + htPrestTva3) + 0.055;
                                            break;
                                        case 0:
                                            totHtProd0 = htProdTva3;
                                            totHtPrest0 = htPrestTva3;
                                            break;
                                        default:
                                            break;
                                    }

                                    //Traitements ventes taux 4
                                    switch (Convert.ToDouble(retourRequete1["TXTAXE4"]))
                                    {
                                        case 20:
                                            totHtProd20 = htProdTva4;
                                            totHtPrest20 = htPrestTva4;
                                            totTva20 = (htProdTva4 + htPrestTva4) + 0.2;
                                            break;
                                        case 10:
                                            totHtProd10 = htProdTva4;
                                            totHtPrest10 = htPrestTva4;
                                            totTva10 = (htProdTva4 + htPrestTva4) + 0.1;
                                            break;
                                        case 5.5:
                                            totHtProd55 = htProdTva4;
                                            totHtPrest55 = htPrestTva4;
                                            totTva55 = (htProdTva4 + htPrestTva4) + 0.055;
                                            break;
                                        case 0:
                                            totHtProd0 = htProdTva4;
                                            totHtPrest0 = htPrestTva4;
                                            break;
                                        default:
                                            break;
                                    }

                                    sw.Write(Convert.ToString(totTva20));
                                    sw.Write(";");
                                    sw.Write(Convert.ToString(totTva10));
                                    sw.Write(";");
                                    sw.Write(Convert.ToString(totTva55));
                                    sw.Write(";");
                                    sw.Write(Convert.ToString(totHtProd20));
                                    sw.Write(";");
                                    sw.Write(Convert.ToString(totHtProd10));
                                    sw.Write(";");
                                    sw.Write(Convert.ToString(totHtProd55));
                                    sw.Write(";");
                                    sw.Write(Convert.ToString(totHtProd0));
                                    sw.Write(";");
                                    sw.Write(Convert.ToString(totHtPrest20));
                                    sw.Write(";");
                                    sw.Write(Convert.ToString(totHtPrest10));
                                    sw.Write(";");
                                    sw.Write(Convert.ToString(totHtPrest55));
                                    sw.Write(";");
                                    sw.Write(Convert.ToString(totHtPrest0));
                                    sw.Write(";");
                                    sw.Write(retourRequete1["TOTTTC"]);
                                }
                                sw.Write(Environment.NewLine);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }
                            finally
                            {
                                // Fermeture de la connexion
                                retourRequete1.Close();
                            }
                        }
                    }
                }
            }
        }
    }
}