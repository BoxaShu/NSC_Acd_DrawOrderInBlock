///https://drive-cad-with-code.blogspot.com/2013/11/draw-order-problem-with-dynamic-block.html
///

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;


using App = Autodesk.AutoCAD.ApplicationServices;
using cad = Autodesk.AutoCAD.ApplicationServices.Application;
using Db = Autodesk.AutoCAD.DatabaseServices;
using Ed = Autodesk.AutoCAD.EditorInput;
using Gem = Autodesk.AutoCAD.Geometry;
using Rtm = Autodesk.AutoCAD.Runtime;


[assembly: Rtm.CommandClass(typeof(NSC_Acd_DrawOrderInBlock.Commands))]

namespace NSC_Acd_DrawOrderInBlock
{

  public class Commands : Rtm.IExtensionApplication
  {

    const string ns = "NSC";



    /// <summary>
    /// Загрузка библиотеки
    /// http://through-the-interface.typepad.com/through_the_interface/2007/03/getting_the_lis.html
    /// </summary>
    #region 
    public void Initialize()
    {
      string assemblyFileFullName = GetType().Assembly.Location;
      string assemblyName = System.IO.Path.GetFileName(
                                           GetType().Assembly.Location);

      // Just get the commands for this assembly
      App.DocumentCollection dm = App.Application.DocumentManager;
      Assembly asm = Assembly.GetExecutingAssembly();


      // Сообщаю о том, что произведена загрузка сборки 
      //и указываю полное имя файла,
      // дабы было видно, откуда она загружена
      App.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage
        (string.Format("\n{0} {1} {2}.\n{3}: {4}\n{5}\n",
                "Assembly", assemblyName, "Loaded",
                "Assembly File:", assemblyFileFullName,
                 "Copyright © ООО 'НСК-Проект' written by Владимир Шульжицкий, 07.2020"));


      //Вывожу список комманд определенных в библиотеке
      App.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage
        ("\nStart list of commands: \n\n");

      string[] cmds = GetCommands(asm, false);
      foreach (string cmd in cmds)
        App.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage
          (cmd + "\n");

      App.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage
        ("\n\nEnd list of commands.\n");


      MakeAutocadTrustMe();
    }

    public void Terminate()
    {
      Console.WriteLine("finish!");
    }

    /// <summary>
    /// Получение списка комманд определенных в сборке
    /// </summary>
    /// <param name="asm"></param>
    /// <param name="markedOnly"></param>
    /// <returns></returns>
    private string[] GetCommands(Assembly asm, bool markedOnly)
    {
      List<string> result = new List<string>();
      object[] objs =
        asm.GetCustomAttributes(typeof(Rtm.CommandClassAttribute), true);
      Type[] tps;
      int numTypes = objs.Length;
      if (numTypes > 0)
      {
        tps = new Type[numTypes];
        for (int i = 0; i < numTypes; i++)
        {
          Rtm.CommandClassAttribute cca =
            objs[i] as Rtm.CommandClassAttribute;
          if (cca != null)
          {
            tps[i] = cca.Type;
          }
        }
      }
      else
      {
        // If we're only looking for specifically
        // marked CommandClasses, then use an
        // empty list
        if (markedOnly)
          tps = new Type[0];
        else
          tps = asm.GetExportedTypes();
      }
      foreach (Type tp in tps)
      {
        MethodInfo[] meths = tp.GetMethods();
        foreach (MethodInfo meth in meths)
        {
          objs = meth.GetCustomAttributes(typeof(Rtm.CommandMethodAttribute), true);
          foreach (object obj in objs)
          {
            Rtm.CommandMethodAttribute attb =
                (Rtm.CommandMethodAttribute)obj; result.Add(attb.GlobalName);
          }
        }
      }
      return result.ToArray();
    }
    #endregion

    protected void MakeAutocadTrustMe()
    {
      try
      {
        string text = Autodesk.AutoCAD.ApplicationServices.Core.Application.GetSystemVariable("TRUSTEDPATHS") as string;
        string directoryName = Path.GetDirectoryName(this.AssemblyPath);
        if (string.IsNullOrEmpty(text) || !text.Contains(directoryName))
        {
          text = directoryName + ";" + text;
          Autodesk.AutoCAD.ApplicationServices.Core.Application.SetSystemVariable("TRUSTEDPATHS", text);
        }
      }
      catch (System.Exception)
      {
      }
    }


    public string AssemblyPath
    {
      get
      {
        return Assembly.GetExecutingAssembly().Location;
      }
    }
    [Rtm.CommandMethod("DOinBlock")]
    public static void DOinBlock()
    {
      App.Document dwg = App.Application.DocumentManager.MdiActiveDocument;
      Ed.Editor ed = dwg.Editor;

      //Ask user to select a block. In real world, I use code
      //to select all targeting blocks by its name, so that
      //user does not have to select block manually
      Ed.PromptEntityOptions opt = new Ed.PromptEntityOptions("\nSelect a block:");
      opt.SetRejectMessage("\nInvalid: not a block.");
      opt.AddAllowedClass(typeof(Db.BlockReference), true);
      Ed.PromptEntityResult res = ed.GetEntity(opt);

      if (res.Status != Ed.PromptStatus.OK) return;

      SetDrawOrderInBlock(dwg, res.ObjectId);
      ed.Regen();

      Autodesk.AutoCAD.Internal.Utils.PostCommandPrompt();
    }

    private static void SetDrawOrderInBlock(App.Document dwg, Db.ObjectId blkId)
    {
      using (Db.Transaction tran =
          dwg.TransactionManager.StartTransaction())
      {
        Db.BlockReference bref = (Db.BlockReference)tran.GetObject(
            blkId, Db.OpenMode.ForRead);

        Db.BlockTableRecord bdef = (Db.BlockTableRecord)tran.GetObject(
            bref.BlockTableRecord, Db.OpenMode.ForWrite);

        Db.DrawOrderTable doTbl = (Db.DrawOrderTable)tran.GetObject(
            bdef.DrawOrderTableId, Db.OpenMode.ForWrite);

        Db.ObjectIdCollection col = new Db.ObjectIdCollection();
        foreach (Db.ObjectId id in bdef)
        {
          if (id.ObjectClass == Rtm.RXObject.GetClass(typeof(Db.Wipeout)))
          {
            col.Add(id);
          }
        }

        if (col.Count > 0)
          doTbl.MoveToBottom(col);


        tran.Commit();
      }
    }

  }
}


