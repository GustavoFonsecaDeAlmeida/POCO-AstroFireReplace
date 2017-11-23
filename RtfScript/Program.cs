using GemBox.Document;
using Independentsoft.Office.Word;
using Spire.Doc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace RtfScript
{
    class Program
    {



        static void Main(string[] args)
        {
            gerar();
            unir();
            replace();
            unir2();
        }


        public static void gerar()
        {
            ComponentInfo.SetLicense("DHPL-8V4A-TLKO-SHIZ");
            



            DocumentModel Capa = new DocumentModel();

            string pathToResources = @"C:\Astrologia\";

            var section = new GemBox.Document.Section(Capa);
            Capa.Sections.Add(section);

            GemBox.Document.Paragraph paragraph = new GemBox.Document.Paragraph(Capa);
            section.Blocks.Add(paragraph);

            Picture picture1 = new Picture(Capa, Path.Combine(pathToResources, "Capa.png"), 5000, 5500, LengthUnit.Pixel);
            paragraph.Inlines.Add(picture1);

            Capa.Save(pathToResources + "Capa.rtf");


            gerar2();


        }

        public static void gerar2()
        {
            ComponentInfo.SetLicense("DHPL-8V4A-TLKO-SHIZ");

            string pathToResources = @"C:\Astrologia\";

            DocumentModel mapa = new DocumentModel();
            var section2 = new GemBox.Document.Section(mapa);
            mapa.Sections.Add(section2);

            GemBox.Document.Paragraph paragraph2 = new GemBox.Document.Paragraph(mapa);
            section2.Blocks.Add(paragraph2);

            Picture picture2 = new Picture(mapa, Path.Combine(pathToResources, "mapa.png"), 5000, 5500, LengthUnit.Pixel);
            paragraph2.Inlines.Add(picture2);

            mapa.Save(pathToResources + "Mapa.rtf");


        }

        public static void unir()
        {

            // If using Professional version, put your serial key below.
            ComponentInfo.SetLicense("DHPL-8V4A-TLKO-SHIZ");


            string pathToResources = @"C:\Astrologia\";

            DocumentModel document = DocumentModel.Load(pathToResources + "Capa.rtf");



            DocumentModel sourceDocument = DocumentModel.Load(Path.Combine(pathToResources, "mapa.rtf"), GemBox.Document.LoadOptions.RtfDefault);

            // Reuse same mapping for importing all sections to improve performance.
            var mapping = new ImportMapping(sourceDocument, document, false);

            // Import all sections from source document.
            foreach (GemBox.Document.Section sourceSection in sourceDocument.Sections)
            {
                GemBox.Document.Section destinationSection = document.Import(sourceSection, true, mapping);
                document.Sections.Add(destinationSection);
            }

            document.Save(pathToResources + "Final.rtf");

        }

        public static void unir2()
        {

            // If using Professional version, put your serial key below.
            ComponentInfo.SetLicense("DHPL-8V4A-TLKO-SHIZ");

            string pathToResources = @"C:\Astrologia\";

            DocumentModel document = DocumentModel.Load(pathToResources + "Final.rtf");



            DocumentModel sourceDocument = DocumentModel.Load(Path.Combine(pathToResources, "DocumentoEditado.rtf"), GemBox.Document.LoadOptions.RtfDefault);

            // Reuse same mapping for importing all sections to improve performance.
            var mapping = new ImportMapping(sourceDocument, document, false);

            // Import all sections from source document.
            foreach (GemBox.Document.Section sourceSection in sourceDocument.Sections)
            {
                GemBox.Document.Section destinationSection = document.Import(sourceSection, true, mapping);
                document.Sections.Add(destinationSection);
            }

            document.Save(pathToResources + "DocumentoEditadoFinal.rtf");

        }

        public static void replace()
        {
            {
                // If using Professional version, put your serial key below.
                ComponentInfo.SetLicense("DHPL-8V4A-TLKO-SHIZ");

                string pathToResources = @"C:\Astrologia\";

                DocumentModel document = DocumentModel.Load(pathToResources + "Sem4.rtf");


                //foreach (ContentRange item in document.Content.Find("ESCOLA CLAUDIA LISBOA DE ASTROLOGIA").Reverse())
                //{
                //    document.Content.Replace("THE MOON IN LIBRA", "A Lua em Libra");
                //    document.Content.Replace("THE MOON", "Lua");
                //    document.Content.Replace("SOLAR FIRE INTERPRETATIONS REPORT", "Gustavo Fonseca de Almeida");
                //    document.Content.Replace("CHART DETAILS", "Detalhes");
                //}



                    //}

                    //for (int i = 0; i < document.Sections.Count(); i++)
                    //{
                    //    var x = document.Sections[i].Content.Find("CHIRON");
                    //    document.Sections[i].Blocks[1].Content.Delete();
                    //    if (x.Count() > 0)
                    //    {
                    //        var a = new ContentRange(document.Sections[i].Blocks.Content.Start, document.Sections[i].Blocks.Content.Start);
                    //           a.Delete();
                    //    }
                    //}
                    var ParagrafosAction = new Dictionary<string, Acao>();


               
                ParagrafosAction.Add("THE MOON\r\n", Acao.MANTER);
                ParagrafosAction.Add("THE MOON IN LIBRA\r\n", Acao.MANTER);
                ParagrafosAction.Add("THE MOON IN THE 3RD HOUSE\r\n", Acao.MANTER);
                ParagrafosAction.Add("ASPECTS OF THE MOON\r\n", Acao.MANTER);
                ParagrafosAction.Add("THE SUN\r\n", Acao.MANTER);
                ParagrafosAction.Add("THE THE SUN IN SCORPIO\r\n", Acao.MANTER);
                ParagrafosAction.Add("THE SUN IN THE 4TH HOUSE\r\n", Acao.MANTER);
                ParagrafosAction.Add("ASPECTS OF THE SUN\r\n", Acao.DELETAR);
                ParagrafosAction.Add("MERCURY\r\n", Acao.MANTER);
                ParagrafosAction.Add("MERCURY IN SAGITTARIUS\r\n", Acao.MANTER);
                ParagrafosAction.Add("MERCURY IN THE 5TH HOUSE\r\n", Acao.MANTER);
                ParagrafosAction.Add("ASPECTS OF MERCURY\r\n", Acao.MANTER);
                ParagrafosAction.Add("VENUS\r\n", Acao.MANTER);
                ParagrafosAction.Add("VENUS IN SCORPIO\r\n", Acao.MANTER);
                ParagrafosAction.Add("VENUS IN THE 4TH HOUSE\r\n", Acao.MANTER);
                ParagrafosAction.Add("ASPECTS OF VENUS\r\n", Acao.MANTER);
                ParagrafosAction.Add("MARS\r\n", Acao.MANTER);
                ParagrafosAction.Add("MARS IN LIBRA\r\n", Acao.MANTER);
                ParagrafosAction.Add("MARS IN THE 3RD HOUSE\r\n", Acao.MANTER);
                ParagrafosAction.Add("ASPECTS OF MARS\r\n", Acao.MANTER);
                ParagrafosAction.Add("JUPITER\r\n", Acao.MANTER);
                ParagrafosAction.Add("JUPITER IN SCORPIO\r\n", Acao.MANTER);
                ParagrafosAction.Add("JUPITER IN THE 4TH HOUSE\r\n", Acao.MANTER);
                ParagrafosAction.Add("ASPECTS OF JUPITER\r\n", Acao.MANTER);
                ParagrafosAction.Add("SATURN\r\n", Acao.MANTER);
                ParagrafosAction.Add("SATURN IN SAGITTARIUS\r\n", Acao.MANTER);
                ParagrafosAction.Add("SATURN IN THE 6TH HOUSE\r\n", Acao.MANTER);
                ParagrafosAction.Add("ASPECTS OF SATURN\r\n", Acao.MANTER);
                ParagrafosAction.Add("SQUARE CHIRON  Orb 1°37' Separating\r\n", Acao.DELETAR);
                ParagrafosAction.Add("URANUS\r\n", Acao.MANTER);
                ParagrafosAction.Add("URANUS IN ARIES\r\n", Acao.MANTER);
                ParagrafosAction.Add("URANUS IN THE 9TH HOUSE\r\n", Acao.MANTER);
                ParagrafosAction.Add("ASPECTS OF URANUS\r\n", Acao.DELETAR);
                ParagrafosAction.Add("NEPTUNE\r\n", Acao.MANTER);
                ParagrafosAction.Add("NEPTUNE IN PISCES\r\n", Acao.MANTER);
                ParagrafosAction.Add("NEPTUNE IN THE 8TH HOUSE\r\n", Acao.MANTER);
                ParagrafosAction.Add("ASPECTS OF NEPTUNE\r\n", Acao.MANTER);
                ParagrafosAction.Add("SEXTILE PLUTO  Orb 5°56' Separating\r\n", Acao.DELETAR);
                ParagrafosAction.Add("PLUTO\r\n", Acao.MANTER);
                ParagrafosAction.Add("PLUTO IN CAPRICORN\r\n", Acao.MANTER);
                ParagrafosAction.Add("PLUTO IN THE 7TH HOUSE\r\n", Acao.MANTER);
                ParagrafosAction.Add("ASPECTS OF PLUTO\r\n", Acao.DELETAR);
                ParagrafosAction.Add("CHIRON\r\n", Acao.DELETAR);
                ParagrafosAction.Add("CHIRON IN THE 8TH HOUSE\r\n", Acao.DELETAR);
                ParagrafosAction.Add("ASPECTS OF CHIRON\r\n", Acao.DELETAR);
                ParagrafosAction.Add("THE NORTH NODE\r\n", Acao.DELETAR);
                ParagrafosAction.Add("THE NORTH NODE IN LEO\r\n", Acao.DELETAR);
                ParagrafosAction.Add("THE NORTH NODE IN THE 1ST HOUSE\r\n", Acao.DELETAR);
                ParagrafosAction.Add("ASPECTS OF THE NORTH NODE\r\n", Acao.DELETAR);
                ParagrafosAction.Add("THE SOUTH NODE\r\n", Acao.DELETAR);
                ParagrafosAction.Add("THE SOUTH NODE IN AQUARIUS\r\n", Acao.DELETAR);
                ParagrafosAction.Add("THE SOUTH NODE IN THE 7TH HOUSE\r\n", Acao.DELETAR);
                ParagrafosAction.Add("ASPECTS OF THE SOUTH NODE\r\n", Acao.DELETAR); 
                ParagrafosAction.Add("THE ASCENDANT\r\n", Acao.MANTER);
                ParagrafosAction.Add("THE ASCENDANT IN CANCER\r\n", Acao.MANTER);
                ParagrafosAction.Add("THE ASCENDANT IN THE 1ST HOUSE\r\n", Acao.DELETAR);
                ParagrafosAction.Add("ASPECTS OF THE ASCENDANT\r\n", Acao.DELETAR);
                ParagrafosAction.Add("THE MIDHEAVEN\r\n", Acao.MANTER);
                ParagrafosAction.Add("THE MIDHEAVEN IN ARIES\r\n", Acao.MANTER);
                ParagrafosAction.Add("THE MIDHEAVEN IN THE 10TH HOUSE\r\n", Acao.DELETAR);
                ParagrafosAction.Add("ASPECTS OF THE MIDHEAVEN\r\n", Acao.DELETAR);
                ParagrafosAction.Add("PT FORTUNE\r\n", Acao.DELETAR);
                ParagrafosAction.Add("PT FORTUNE IN GEMINI\r\n", Acao.DELETAR);
                ParagrafosAction.Add("PT FORTUNE IN THE 11TH HOUSE\r\n", Acao.DELETAR);
                ParagrafosAction.Add("ASPECTS OF PT FORTUNE\r\n", Acao.DELETAR);
                ParagrafosAction.Add("THE BLACK MOON\r\n", Acao.DELETAR);
                ParagrafosAction.Add("THE BLACK MOON IN CAPRICORN\r\n", Acao.DELETAR);
                ParagrafosAction.Add("THE BLACK MOON IN THE 6TH HOUSE\r\n", Acao.DELETAR);
                ParagrafosAction.Add("ASPECTS OF THE BLACK MOON\r\n", Acao.DELETAR);

                ParagrafosAction.ToList();


                var Paragrafos = new List<GemBox.Document.Paragraph>();
                var paragrafosToDelete = new List<GemBox.Document.Paragraph>();
                var resultado = "";
                var indexItem = -1;
                var indexitem2 = 0;
                var primeiroAdeletar = 0;
                var finalADeletar = 0;
      
                var k = new List<GemBox.Document.Paragraph>();

                var indexparagrafoatual = 0;
                var IndexProximoParagrafo = 0;

                // Get content from each paragraph
                foreach (GemBox.Document.Paragraph paragraph in document.GetChildElements(true, ElementType.Paragraph))
                {
                    Paragrafos.Add(paragraph);
                }


                foreach (var item in Paragrafos)
                {
                    indexItem++;
                    foreach (var item2 in ParagrafosAction)
                    {
                        if (item.Content.ToString() == item2.Key.ToString())
                        {
                            if (item2.Value == Acao.DELETAR)
                            {
                                ProximoElementoDicionario(ParagrafosAction, item.Content.ToString(), out resultado);
                                if (resultado == item.Content.ToString())
                                {
                                    k.AddRange(Paragrafos.GetRange(indexItem, Paragrafos.Count() - indexItem));
                                    paragrafosToDelete.AddRange(k);
                                }
                                else
                                {
                                    ContarAteOProximo(Paragrafos, item.Content.ToString(), resultado.ToString(), out indexparagrafoatual, out IndexProximoParagrafo);
                                    k.AddRange(Paragrafos.GetRange(indexparagrafoatual, IndexProximoParagrafo - indexparagrafoatual));
                                    paragrafosToDelete.AddRange(k);
                                }
                            }
                        }
                    }
                }

                //foreach (var item in Paragrafos)
                //{
                //    indexItem++;
                //    foreach (var item2 in ParagrafosAction)
                //    {

                //        if (item.Content.ToString() == item2.Key.ToString())
                //        {
                //            if (item2.Value == Acao.DELETAR)
                //            {
                //                primeiroAdeletar = indexItem;
                //                ProximoElementoDicionario(ParagrafosAction, item.Content.ToString(), out resultado);
                //                possuiuprimeirodelete = true;
                //            }



                //        }

                //        if (item.Content.ToString() == resultado.ToString())
                //        {
                //            if (item2.Value == Acao.DELETAR)
                //            {
                //                finalADeletar = indexItem - 1;
                //                possuiultimodelete = true;
                //            }



                //        }
                //        if (primeiroAdeletar != 0 && finalADeletar != 0 && possuiultimodelete == true && possuiultimodelete == true)
                //        {
                //            if (finalADeletar < primeiroAdeletar)
                //            {
                //                finalADeletar = Paragrafos.Count();
                //            }
                //            k.AddRange(Paragrafos.GetRange(primeiroAdeletar, finalADeletar - primeiroAdeletar));
                //            paragrafosToDelete.AddRange(k);
                //            primeiroAdeletar = 0;
                //            finalADeletar = 0;
                //            possuiultimodelete = false;
                //            possuiuprimeirodelete = false;

                //        }
                //    }

                //}


                foreach (var item in paragrafosToDelete)
                {

                    item.Content.Delete();
                }

                foreach (ContentRange item in document.Content.Find("SOLAR FIRE INTERPRETATIONS REPORT").Reverse())
                {

                    document.Content.Replace("THE MIDHEAVEN IN ARIES", "");
                    document.Content.Replace("THE MIDHEAVEN", "");
                    document.Content.Replace("THE ASCENDANT IN CANCER", "");
                    document.Content.Replace("THE ASCENDANT", "");
                    document.Content.Replace("PLUTO IN THE 7TH HOUSE", "");
                    document.Content.Replace("PLUTO IN CAPRICORN", "");
                    document.Content.Replace("PLUTO", "");
                    document.Content.Replace("ASPECTS OF NEPTUNE", "");
                    document.Content.Replace("NEPTUNE IN THE 8TH HOUSE", "");
                    document.Content.Replace("NEPTUNE IN PISCES", "");
                    document.Content.Replace("NEPTUNE", "");
                    document.Content.Replace("URANUS IN THE 9TH HOUSE", "");
                    document.Content.Replace("URANUS IN ARIES", "");
                    document.Content.Replace("ASPECTS OF SATURN", "");
                    document.Content.Replace("SATURN IN THE 6TH HOUSE", "");
                    document.Content.Replace("SATURN IN SAGITTARIUS", "");
                    document.Content.Replace("SATURN", "");
                    document.Content.Replace("TRINE NEPTUNE  Orb 3°33' Applying", "");
                    document.Content.Replace("ASPECTS OF JUPITER", "");
                    document.Content.Replace("JUPITER IN THE 4TH HOUSE", "");
                    document.Content.Replace("JUPITER IN SCORPIO", "");
                    document.Content.Replace("JUPITER", "");
                    document.Content.Replace("SQUARE PLUTO  Orb 2°06' Applying", "");
                    document.Content.Replace("ASPECTS OF MARS", "");
                    document.Content.Replace("MARS IN THE 3RD HOUSE", "");
                    document.Content.Replace("MARS IN LIBRA", "");
                    document.Content.Replace("MARS", "");
                    document.Content.Replace("TRINE NEPTUNE  Orb 0°45' Applying", "");
                    document.Content.Replace("CONJUNCTION JUPITER  Orb 2°48' Separating", "");
                    document.Content.Replace("ASPECTS OF VENUS", "");
                    document.Content.Replace("VENUS IN THE 4TH HOUSE", "");
                    document.Content.Replace("VENUS IN SCORPIO", "");
                    document.Content.Replace("VENUS", "");
                    document.Content.Replace("SQUARE NEPTUNE  Orb 2°47' Separating", "");
                    document.Content.Replace("SEXTILE MARS  Orb 1°02' Applying", "");
                    document.Content.Replace("ASPECTS OF MERCURY", "");
                    document.Content.Replace("MERCURY IN THE 5TH HOUSE", "");
                    document.Content.Replace("MERCURY IN SAGITTARIUS", "");
                    document.Content.Replace("MERCURY", "");
                    document.Content.Replace("THE SUN IN THE 4TH HOUSE", "");
                    document.Content.Replace("THE SUN IN SCORPIO", "");
                    document.Content.Replace("THE SUN", "");
                    document.Content.Replace("OPPOSITION URANUS  Orb 0°42' Separating", "");
                    document.Content.Replace("SEXTILE SATURN  Orb 0°03' Separating", "");
                    document.Content.Replace("ASPECTS OF THE MOON", "");
                    document.Content.Replace("THE MOON IN THE 3RD HOUSE", "");
                    document.Content.Replace("THE MOON IN LIBRA", "");
                    document.Content.Replace("SEXTILE   Orb 0°03' Separating", "");
                    document.Content.Replace("SEXTILE   Orb 1°02' Applying", "");
                    document.Content.Replace("SQUARE   Orb 2°47' Separating", "");
                    document.Content.Replace("CONJUNCTION Orb 2°48' Separating", "");
                    document.Content.Replace("CONJUNCTION   Orb 2°48' Separating", "");
                    document.Content.Replace("SQUARE   Orb 2°06' Applying", "");
                    document.Content.Replace("TRINE   Orb 0°45' Applying", "");
                    document.Content.Replace("TRINE URANUS  Orb 0°38' Separating", "");
                    document.Content.Replace("URANUS", "");
                    document.Content.Replace("THE MOON", "");
                    document.Content.Replace("TRINE   Orb 3°33' Applying", "");




                    //document.Content.Replace("THE MOON IN LIBRA", "A Lua em Libra");
                    //document.Content.Replace("THE MOON", "Lua");
                    //document.Content.Replace("SOLAR FIRE INTERPRETATIONS REPORT", "Gustavo Fonseca de Almeida");
                    //document.Content.Replace("CHART DETAILS", "Detalhes");
                }



                document.Save(pathToResources + "DocumentoEditado.rtf");
            }
        }

        public static void delete(DocumentModel document, string toDelete)
        {

            var blocks = document.Sections.ToList();
            foreach (var item in blocks)
            {
                foreach (var item2 in item.Blocks)
                {
                    var a = new ContentRange(item2.Content.Start, item2.Content.End);
                    a.Delete();
                }

            }
        }

        public static void ProximoElementoDicionario(Dictionary<string, Acao> Lista, string valorAtual, out string resultado)
        {
            resultado = valorAtual.ToString();
            for (int i = 0; i < Lista.Count(); i++)
            {
                var x = Lista.ElementAt(i);

                if (x.Key.ToString() == valorAtual)
                {
                    if (i == Lista.Count() - 1)
                    {
                        var y = Lista.ElementAt(i);
                        resultado = y.Key.ToString();
                    }
                    else
                    {
                        var y = Lista.ElementAt(i + 1);
                        resultado = y.Key.ToString();
                    }
                     
                }

            }

            //resultado = valorAtual.ToString();
            //var achou = false;
            //var vaiembora = false;
            //foreach (var item in Lista)
            //{
            //    if (vaiembora == true)
            //    {
            //        continue;
            //    }
            //    if (achou == false)
            //    {
            //        if (item.Key.ToString() == valorAtual)
            //        {
            //            achou = true;
            //        }
            //    }
            //    else
            //    {
            //        if (achou == true)
            //        {
            //            resultado = item.Key;
            //            vaiembora = true;
            //        }

            //    }
            //}

        }

        public static void ContarAteOProximo(List<GemBox.Document.Paragraph> Paragrafos, string paragrafoAtual , string proximoParagrafo , out int IndexparagrafoAtual , out int IndexProximoParagrafo) {
            var index = -1;
            IndexparagrafoAtual = 0;
            IndexProximoParagrafo = 0;

            foreach (var item in Paragrafos)
            {
                index++;
                if (item.Content.ToString() == paragrafoAtual.ToString())
                {
                    IndexparagrafoAtual = index;
                }
                if (item.Content.ToString() == proximoParagrafo.ToString())
                {
                    IndexProximoParagrafo = index;
                }

            }

        }

        
    }
}







