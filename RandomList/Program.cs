using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;

namespace RandomList
{
    class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                try
                {
                    string currentDirectory = Directory.GetCurrentDirectory();

                    string studentFileName = "estudiantes.txt";
                    Console.Write("Nombre del archivo de Estudiantes: ");
                    studentFileName = Console.ReadLine();
                    string studentPath = currentDirectory + "\\" + studentFileName;
                    
                    //Comprobar si tiene el archivo de estudiantes
                    if (File.Exists(studentPath))  
                    {
                        while (true)
                        {
                            string themeFileName = "temas.txt";
                            Console.Write("Nombe del archivo de Temas: ");
                            themeFileName = Console.ReadLine();
                            string themePath = currentDirectory + "\\" + themeFileName;

                            //Comprobar si tiene el archivo de temas
                            if (File.Exists(themePath)) 
                            {

                                Console.Write("Inserte la cantidad de estudiantes por grupos: ");
                                int cantEstGroup = int.Parse(Console.ReadLine()); 

                                //Inicializando listas para hacer los procesos
                                List<Grupo> groups = new List<Grupo>();
                                List<Estudiante> students = new List<Estudiante>();
                                List<Tema> themes = new List<Tema>();

                                using (StreamReader reader = new StreamReader(studentPath))
                                {
                                    string line;
                                    Estudiante student;
                                    while ((line = reader.ReadLine()) != null)
                                    {                         
                                        /*Para cada linea de estudiante en el archivo crea
                                        un objeto clase Estudiante y lo introduce a la lista*/
                                        
                                        student = new Estudiante();
                                        student.FullName = line;
                                        students.Add(student);
                                    }
                                }

                                using (StreamReader reader = new StreamReader(themePath))
                                {
                                    string line;
                                    Tema theme;
                                    while ((line = reader.ReadLine()) != null)
                                    {
                                        /*Para cada linea de temas en el archivo crea
                                        un objeto clase Tema y lo introduce a la lista*/
                                        theme = new Tema();
                                        theme.Name = line;
                                        themes.Add(theme);
                                    }
                                }

                                if (students.Count >= cantEstGroup && themes.Count >= cantEstGroup)
                                {
                                    int groupsCount = students.Count / cantEstGroup;    // Calcula cantidad de grupos
                                    int rest = students.Count % cantEstGroup;           // Calcula cuantos estudiantes restan
                                    int themesCount = themes.Count / groupsCount;       // Calcula cantidad de temas por grupo
                                    int restThemes = themes.Count % groupsCount;        // Calcula cantidad de temas restantes

                                    //Se realizan validaciones según requerimientos
                                    if (students.Count >= groupsCount && themes.Count >= groupsCount)
                                    {
                                        Random random = new Random();
                                        Grupo group;
                                        List<Estudiante> studentList;
                                        List<Tema> themeList;
                                        for (int a = 0; a < groupsCount; a++)
                                        {
                                            // Inicializa un nuevo grupo en cada iteracion
                                            group = new Grupo();
                                            group.Nro = a + 1;
                                            studentList = new List<Estudiante>();
                                            themeList = new List<Tema>();
                                            for (int b = 0; b < cantEstGroup; b++)
                                            {
                                                /* Toma un estudiante aleatorio de la lista de estudiantes 
                                                 lo agrega a la lista de estudiantes de ese grupo y luego lo
                                                 remueve de la lista original de estudiantes */
                                                int i = random.Next(0, students.Count);
                                                studentList.Add(students[i]);
                                                students.RemoveAt(i);
                                            }
                                            for (int c = 0; c < themesCount; c++)
                                            {
                                                /* Toma un tema aleatorio de la lista de temas 
                                                 lo agrega a la lista de temas de ese grupo y luego lo
                                                 remueve de la lista original de temas */
                                                int i = random.Next(0, themes.Count);
                                                themeList.Add(themes[i]);
                                                themes.RemoveAt(i);
                                            }
                                            group.Estudiantes = studentList;
                                            group.Temas = themeList;
                                            groups.Add(group);
                                        }
                                        List<int> counts = new List<int>();
                                        bool valid = false;
                                        while (rest > 0)
                                        {
                                            valid = false;
                                            int i = random.Next(0, groups.Count);
                                            if (counts.Count > 0 && groups.Count > 1)
                                            {
                                                while (!valid)
                                                {
                                                    i = random.Next(0, groups.Count);
                                                    foreach (var item in counts)
                                                    {
                                                        valid = true;
                                                        if (i == item)
                                                        {
                                                            valid = false;
                                                            break;
                                                        }
                                                    }
                                                }
                                            } 
                                            // Comprueba si el grupo al que se quiere agregar es el indicado
                                            counts.Add(i);
                                            if (counts.Count == groups.Count)
                                                counts.Clear();

                                            int x = random.Next(0, students.Count);
                                            List<Estudiante> studentsTemp = groups[i].Estudiantes;//Copia la lista
                                            Estudiante studentTemp = students[x];

                                            studentsTemp.Add(studentTemp);                        //Agrega el estudiante a la lista temporal
                                            groups[i].Estudiantes = studentsTemp;                 //Sustituye la lista
                                            students.RemoveAt(x);                                 //Remueve ese estudiante restante de la lista de estudiantes
                                            rest--;                                               //Continua haciendo ese proceso hasta que ya no haya ESTUDIANTES restantes
                                        }
                                        counts.Clear();                                           //REINICIA Counts para hacer el mismo proceso con los temas
                                        while (restThemes > 0)
                                        {
                                            int i = random.Next(0, groups.Count);
                                            valid = false;
                                            if (counts.Count > 0 && groups.Count > 1)
                                            {
                                                while (!valid)
                                                {
                                                    i = random.Next(0, groups.Count);
                                                    foreach (var item in counts)
                                                    {
                                                        valid = true;
                                                        if (i == item)
                                                        {
                                                            valid = false;
                                                            break;
                                                        }
                                                    }
                                                }
                                            } // Comprueba si el grupo al que se quiere agregar es el indicado
                                            counts.Add(i);
                                            if (counts.Count == groups.Count)
                                                counts.Clear();
                                            int x = random.Next(0, themes.Count);
                                            List<Tema> themesTemp = groups[i].Temas;
                                            Tema themeTemp = themes[x];

                                            themesTemp.Add(themeTemp);    //Agrega el estudiante a la lista temporal
                                            groups[i].Temas = themesTemp; //Sustituye la lista
                                            themes.RemoveAt(x);           //Remueve ese tema restante de la lista de temas
                                            restThemes--;                 //Continua haciendo ese proceso hasta que ya no haya TEMAS restantes
                                        }
                                        string fileName = $"Resultado-{DateTime.Now.ToString("yyyyMMddHHmmss")}.txt";
                                        string filePath = currentDirectory + "\\Resultados\\" + fileName;
                                        if(!Directory.Exists(currentDirectory + "\\Resultados"))
                                        {
                                            Directory.CreateDirectory(currentDirectory + "\\Resultados");
                                        }
                                        /*if (File.Exists(filePath))
                                        {
                                            string[] files = Directory.GetFileSystemEntries(currentDirectory + "\\Resultados", "*.txt");
                                            for(int i = 0; i < files.Count(); i++)
                                            {
                                                files[i] = files[i].Split('\\')[files[i].Split('\\').Count() - 1];
                                            }
                                            
                                            string[] divition = files[files.Count() - 1].Split('-');
                                            divition[1] = divition[1].Split('.')[0];
                                            int numberFile = int.Parse(divition[1]);
                                            numberFile += 1;
                                            filePath = currentDirectory + $"\\Resultados\\Resultado-{numberFile}.txt";
                                        }*/
                                        using (StreamWriter sw = File.CreateText(filePath))
                                        {

                                            //Escribe en el txt de resultados los grupos creados y sus integrantes y temas

                                            sw.WriteLine("==========================================================================");
                                            foreach (var grp in groups)
                                            {
                                                
                                                sw.WriteLine($"Grupo #{grp.Nro}");
                                                sw.WriteLine("----------------------------------------------------------------------");
                                                int i = 1;
                                                foreach (var student in grp.Estudiantes)
                                                {
                                                    sw.WriteLine($"{i}. {student.FullName}");
                                                    i++;
                                                }
                                                sw.WriteLine("----------------------------------------------------------------------");
                                                i = 1;
                                                foreach (var theme in grp.Temas)
                                                {
                                                    sw.WriteLine($"Tema #{i}: {theme.Name}.");
                                                    i++;
                                                }
                                                sw.WriteLine("\n======================================================================");
                                            }
                                        }

                                        using (StreamReader sr = new StreamReader(filePath))
                                        {
                                            // Lee cada linea del archivo creado y lo muestra en la consola
                                            string line;
                                            while ((line = sr.ReadLine()) != null)
                                            {
                                                Console.WriteLine(line);
                                            }
                                        }
                                    }
                                    else
                                        Console.WriteLine("Hay mas grupos que estudiantes.");
                                }
                                else
                                    Console.WriteLine("Hay menos estudiantes o temas que el minimo requerido.");
                                break;
                            }
                            else
                                Console.WriteLine("Archivo no existe.");
                        }
                    }
                    else
                        Console.WriteLine("Archivo no existe");
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                Console.ReadKey();
            }
        }

    }
}
