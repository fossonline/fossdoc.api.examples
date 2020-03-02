using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Foss.FossDoc.ApplicationServer;
using DS;
using Converters = Foss.FossDoc.ApplicationServer.Converters;
using System.IO;
using Foss.FossDoc.ApplicationServer.ObjectDataManagment;
using Foss.FossDoc.ApplicationServer.Messaging;
using Foss.FossDoc.ExternalModules.BusinessLogic.Utils;
using Foss.FossDoc.ApplicationServer.IO;


namespace document.outgoing.letter
{
	/// <summary>
	/// Приклад створення документу "Вихідний лист", додавання файлів та ЕЦП
	/// </summary>
	class Program
	{
		static void Main(string[] args)
		{
			try
			{

				string connectionString = "URL=corbaloc::1.2@andrewsalko:11301/AccessRoot;Login=Деловод;Password=123;AuthenticationAlgorithm=FossDoc;";

				using (ISession session = (ISession)Foss.FossDoc.ApplicationServer.Connection.Connector.Connect(connectionString))
				{
					OID userID = session.AccessControlManager.RoleManager.ActAsUserOID;	//поточний користувач

					var userMngr = session.AccessControlManager.UserManagment;
					bool isUserRegistrator = userMngr.IsUserInGroup(userID, Foss.FossDoc.ExternalModules.EDMS.Schema.Groups.Registers);

					if (!isUserRegistrator)
					{
						throw new ApplicationException("Для створення та реєстрації Вихідного листа користувач має бути у групі 'Реєстратори'");
					}

					//отримаємо папку "Вихідні" у поточного користувача
					using (var currentUserDepartment = (Foss.FossDoc.ExternalModules.EDMS.Interfaces.IDepartment)session.ExternalModulesManager.GetExternalModuleInterface(Foss.FossDoc.ExternalModules.EDMS.Schema.ExternalModule.Name, typeof(Foss.FossDoc.ExternalModules.EDMS.Interfaces.IDepartment).FullName))
					{
						OID folderID = currentUserDepartment.OutboxFolder;
						//якщо папку не буде знайдено, отримаєте виключення

						//створення документу

						DateTime documentRegTime = DateTime.Now.Date;
						string content = "Короткий зміст документа";

						//вам необхідно знайти потрібного кореспондента, у довіднику
						//можна шукати за будь-яким полем, приклад нижче, дивись класс  OrganizationFinder щодо деталей

						string searchKey = "F0000004";
						var obj = session.ObjectDataManager;
						OrganizationFinder organizationFinder = new OrganizationFinder(obj);
						OID corrID = organizationFinder.FindOrganization(searchKey);

						if (corrID.IsEmptyOrUnspecified())
						{
							throw new ApplicationException("Кореспондента не знайдено");
						}

						OutgoingLetterWithAutoRegistration letter = new OutgoingLetterWithAutoRegistration(session, folderID, documentRegTime, content, new OID[] { corrID });

						//тепер додаємо вкладені файл
						string fileName = @"C:\!\Document.docx";
						string fileNameSignature = @"C:\!\Document.docx.p7s";

						Attach attach = new Attach(fileName);

						//важливо - класс Attach генерує OID для файлу самостійно, тому ми нижче використаємо його , щоб додати ЕЦП

						letter.Attaches = new Attach[] { attach };
						
						OID docID=letter.Create();

						//тільки зараз об'єкт-документ та файл вже існують у базі. Тепер додаємо ЕЦП-об'єкт
						DigitalSignature digitalSignature = new DigitalSignature(obj);

						//ім'я ЕЦП-запису таке саме як файлу
						string fileNameOnly = Path.GetFileName(fileName);

						digitalSignature.CreateSignatureRecord(attach.ID, fileNameOnly, userID, fileNameSignature);
					}

				}//using
			}
			catch (Exception ex)
			{
				_Log(ex.ToString());
			}
		}

		static void _Log(string message, params object[] args)
		{
			Console.WriteLine(message, args);
		}


	}
}
