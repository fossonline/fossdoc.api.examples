using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using DS;
using Foss.FossDoc.ApplicationServer;
using Converters = Foss.FossDoc.ApplicationServer.Converters;

namespace document.incoming.processing
{
	class Program
	{
		/// <summary>
		/// Папка Вхідні (в підрозділі, де зберігаються вхідні листи, канцелярія)
		/// Замініть на свою папку (бібл. Сервер \ Папки \ Домашні папки підрозділів \ (Ваш підрозділ) \ Вхідні --- взяти ідентифікатор з "Об'єкт")
		/// </summary>
		static readonly OID IncomingLettersFolderID = Converters.OID.FromString("0000000097916CD706CA4F26A84C60E0BE97E15D");

		/// <summary>
		/// Строка підключення до сервера FossDoc
		/// (змініть параметри пароль та машину)
		/// </summary>
		const string ConnectionString = "URL=corbaloc::1.2@localhost:11301/AccessRoot;Login=Administrator;Password=123;AuthenticationAlgorithm=FossDoc;";

		static void Main(string[] args)
		{
			ISession session = null;
			try
			{
				Console.WriteLine("Incoming documents analyser utility");

				session = (ISession)Foss.FossDoc.ApplicationServer.Connection.Connector.Connect(ConnectionString);

				string folderForDocuments = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

				//Утіліта розрахована на періодичний запуск. Кожного разу вона буде брати до обробки лише нові документи
				//(це визначається пошуком, у нових не буде спеціального прихованого поля, яке ця утіліта проставить в документи які оброблялись).
				//Таким чином навіть якщо в папці багато документів, утіліта не буде постійно "захоплювати" їх всі, а обробляє лише "нові" (де немає властивості-позначки).

				if (!session.ObjectDataManager.GetExistance(IncomingLettersFolderID))
				{
					Console.WriteLine($"Folder not exists: {IncomingLettersFolderID.ToStringRepresentation()}");
					return;
				}

				IncomingAnalyser incomingAnalyser = new IncomingAnalyser(session, IncomingLettersFolderID, folderForDocuments);
				int processed = incomingAnalyser.Perform();

				Console.WriteLine($"Done, processed: {processed}");
			}
			catch (Exception ex)
			{
				Console.WriteLine($"Error: {ex}");
			}
			finally
			{
				if (session != null)
				{
					try
					{
						session.Dispose();
					}
					catch (Exception)
					{
					}
				}
			}
		}
	}
}
