using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Foss.FossDoc.ApplicationServer;
using Foss.FossDoc.ApplicationServer.Converters.Restrictions;
using DS;
using Converters = Foss.FossDoc.ApplicationServer.Converters;
using System.IO;
using Foss.FossDoc.ApplicationServer.ObjectDataManagment;
using Foss.FossDoc.ApplicationServer.IO;

namespace load.@new.emails
{
	/// <summary>
	/// Демонструє підключення до сервера (від імені користувача), перевірку чи є нові листи (email), та завантажує файли з листа та деякі властивості.
	/// </summary>
	class Program
	{
		static void Main(string[] args)
		{
			try
			{

				//Строка підключення до сервера. Логін - dilovod
				//Змініть localhost на і'мя машини де у вас сервер.
				//Якщо у вас SSL-з'єднання, формат строки інший: "iiop-ssl://localhost:11301/AccessRoot;Login=dilovod;Password=123;AuthenticationAlgorithm=FossDoc;ClientAuthentication=Foss.FossDoc.ApplicationServer.Connection.ClientSSLAuthentication, Foss.FossDoc.ApplicationServer.Connection;"
				string connectionString = "URL=corbaloc::1.2@andrewsalko:11301/AccessRoot;Login=Деловод;Password=123;AuthenticationAlgorithm=FossDoc;";

				//Якщо у вас Windows-аутентифікація (ActiveDirectory) то логін та пароль передавати не потрібно:
				//string connectionString = "URL=corbaloc::1.2@andrewsalko:11301/AccessRoot;AuthenticationAlgorithm=Windows;";

				//Ви можете зберігати сессію достатньо довго, але не забувайте про Dispose
				using (ISession session = (ISession)Foss.FossDoc.ApplicationServer.Connection.Connector.Connect(connectionString))
				{
					//менеджер для роботи з об'єктами (через нього ви отримаєте будь-які властивості, наприклад Тема листа, відправник, наявність файлів та інше)
					var obj = session.ObjectDataManager;

					//Сервер сам завантажує та зберігає листи у базі. Наш додаток лише перевіряє, чи є у папці Вхідні нові листи (які поточний користувач не читав)
					//Нам потрібно знати папку Вхідні (її ідентифікатор) для поточного користувача
					OID inboxID = session.MessagingManager.GetSpecialMessagingFolder(Foss.FossDoc.ApplicationServer.Messaging.SpecialMessagingFolder.Inbox);

					//структура DS.OID схожа на Guid, але має додаткові 4 байти. усі об'єкти у FossDoc мають свій унікальний ідентифікатор

					//Нижче ми готуємо умову для пошуку (restrict).
					//Умова "Прочитано" НЕ рівно true. Але це спрацює лише коли властивість встановлено (а для нового листа його якраз не буде)
					PropertyRestrictionHelper propRead = new PropertyRestrictionHelper(Foss.FossDoc.ApplicationServer.ObjectDataManagment.Schema.PropertyTags.ObjectReaded, true, DS.relopNE.ConstVal);
					TableRestrictionHelper tblNotReadLetters = new TableRestrictionHelper(DS.resProperty.ConstVal, propRead);

					//а ось цей блок відбирає "властивість Прочитано не знайдено"
					ExistRestrictionHelper readExists = new ExistRestrictionHelper(Foss.FossDoc.ApplicationServer.ObjectDataManagment.Schema.PropertyTags.ObjectReaded);
					TableRestrictionHelper tblReadExists = new TableRestrictionHelper(DS.resExist.ConstVal, readExists);
					TableRestrictionHelper tblNot = new TableRestrictionHelper(DS.resNOT.ConstVal, tblReadExists);

					//Комплексна умова OR містить дві умови: "Прочитано НЕ є true" та "Прочитано не існує" 
					TableRestrictionHelper tblMain = new TableRestrictionHelper(DS.resOR.ConstVal, tblNot, tblNotReadLetters);

					//Метод GetChildren отримує дочірні обєкти в батьківському об'єкті за вказаним тегом.
					//Тег - можна вважати іменем властивості, хоча це структура з 4 байтів, наприклад: 0x360F1F00 - це тег "Документи в папці".
					//Наприклад у папки є гілки "Документи в папці" та "Папки в папці". У листа - "Файли".
					//наведений приклад GetChildren отримує листи (ідентифікатори) з папки Вхідні за достатньо популярним тегом "Документи в папці", із застосуванням фільтру tblMain.
					//У папці може бути тисячі листів, а нам потрібні тількі нові (непрочитані).

					OID[] messagesIDs = obj.GetChildren(new OID[] { inboxID }, new TPropertyTag[] { Foss.FossDoc.ApplicationServer.Messaging.Schema.MessagesFolder.PR_CONTAINER_CONTENTS }, tblMain);
					if (messagesIDs == null || messagesIDs.Length == 0)
					{
						_Log("Нових листів не знайдено");
						return;//нема нових email
					}

					string appDir = new Uri(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)).LocalPath;

					for (int i = 0; i < messagesIDs.Length; i++)
					{
						OID msgID = messagesIDs[i];

						//тут ми просто виводимо ідентифікатор листа, його можна побачити на вкладці Об'єкт.
						_Log("Лист {0}", msgID.ToStringRepresentation());

						//Тепер ми можемо отримати властивості листа. Багато з них можна побачити на вкладці Обєкт, але контейнерні властивості треба дивитися у дереві (або окремих вкладках)
						//Це стосується "Файли" або "Отримувачі".
						//А поки що ми запитаємо лише дві властовості - Тема та Адреса відправника.

						//tags - це масив властовостей, які ми хочемо запитати в сервера. Багато з них можна знайти в інтерфейсах сервера, або ввести власноруч якщо скопіювати тег з адміністрування.
						//На вкладці Об'єкт, угорі активуйте показ тегів та унікальних імен (Show tags, Show unique names)
						//var testTag = Converters.PropertyTag.FromUInt32(0x0C1F001E);    //це те ж саме що й PR_SENDER_EMAIL_ADDRESS

						TPropertyTag[] tags = new TPropertyTag[]
						{
						Foss.FossDoc.ApplicationServer.Messaging.Schema.Message.PR_SUBJECT,
						Foss.FossDoc.ApplicationServer.Messaging.Schema.Message.PR_SENDER_EMAIL_ADDRESS
						};

						var props = obj.GetProperties(msgID, tags);

						if (props == null)
						{
							//Важливо завжди перевіряти результат GetProperties на null. Якщо у вас немає доступу на читання обєкта або його не існує у базі, ви отримаєте null.
							_Log("Об'єкт не існує або немає доступу на читання: {0}", msgID.ToStringRepresentation());
							continue;
						}

						//Ми запитали дві властивості, та у результаті сервер збереже порядок - першим йде PR_SUBJECT (Тема листа), а потім адреса відправника
						if (props[0].PropertyTag.IsEquals(Foss.FossDoc.ApplicationServer.Messaging.Schema.Message.PR_SUBJECT))
						{
							//Сервер повертає спец.структуру, та в неї перевірте тег - він повинен співпасти. Якщо ні - властивість не існує, та звертатися до значення не можна.
							string subject = props[0].Value.GetstrVal();
							_Log("Тема:{0}", subject);
						}

						if (props[1].PropertyTag.IsEquals(Foss.FossDoc.ApplicationServer.Messaging.Schema.Message.PR_SENDER_EMAIL_ADDRESS))
						{
							string senderAddr = props[1].Value.GetstrVal();
							_Log("Адреса відправника:{0}", senderAddr);
						}

						//Ми створимо окрему папку для кожного листа, за його ідентифікатором:
						string destFolderPath = Path.Combine(appDir, msgID.ToStringRepresentation());
						if (Directory.Exists(destFolderPath))
						{
							//якщо папка вже існує, ми її видалимо
							Directory.Delete(destFolderPath, true);
						}

						Directory.CreateDirectory(destFolderPath);

						//У листа можуть бути файли, але не обов'язково. Метод нижче це проаналізує та завантажить файли на диск
						_DownloadAttaches(obj, msgID, destFolderPath);

						//Якщо необхідно "Прочитати" лист як звичайній користувач - відкрийте цей виклик:
						//session.MessagingManager.SetReadFlag(new OID[] { msgID }, true);
					}

					_Log("Виконано");

				}//using
			}
			catch (Exception ex)
			{
				_Log(ex.ToString());
			}
		}

		static string _DownloadAttach(IObjectDataManager obj, OID attachID, string destinationFolderPath, int blockSize)
		{
			ObjectProperty[] props = obj.GetProperties(attachID,
				Foss.FossDoc.ApplicationServer.Messaging.Schema.Attachment.PR_ATTACH_LONG_FILENAME,
				Foss.FossDoc.ApplicationServer.Messaging.Schema.Attachment.PR_ATTACH_DATA_BIN);

			if (props == null || props.Length == 0 || !props[0].PropertyTag.IsEquals(Foss.FossDoc.ApplicationServer.Messaging.Schema.Attachment.PR_ATTACH_LONG_FILENAME))
			{
				throw new ApplicationException(string.Format("Файл {0} не має імені або не існує", attachID.ToStringRepresentation()));
			}

			string fileName = props[0].Value.GetstrVal();//ім'я файлу
			string destFullFileName = Path.Combine(destinationFolderPath, fileName);

			//якщо тіло файлу невелике, його може повернути GetProperties
			if (props[1].PropertyTag.IsEquals(Foss.FossDoc.ApplicationServer.Messaging.Schema.Attachment.PR_ATTACH_DATA_BIN))
			{
				byte[] smallBody = props[1].Value.GetbinVal();
				File.WriteAllBytes(destFullFileName, smallBody);
			}
			else
			{
				//тіла файлу немає, або воно велике, його потрібно завантажувати через потік
				if ((props[1].Value.Discriminator == ptError.ConstVal) && (props[1].Value.GetlErr() == PropertyErrors.TooBig))
				{
					int buffer = 1024 * 256;//по 256 Кб блок
					if (blockSize > 0)
					{
						buffer = blockSize;//або нам ззовні вказали розмір блоку
					}

					//завантажуємо тіло порціями:
					using (var destStream = File.Create(destFullFileName))
					{
						try
						{
							using (var srcStream = StreamAccessOptimizer.WrapStream(obj.UpdateBinaryStream(attachID, Foss.FossDoc.ApplicationServer.Messaging.Schema.Attachment.PR_ATTACH_DATA_BIN, BinaryStreamAccessMode.Read)))
							{
								Converters.StreamCopy.Perform(srcStream, destStream, buffer);
							}
						}
						catch (DS.PropertyNotExist)
						{
							//немає властивості - файл нульової довжини
						}
					}//using
				}
				else
				{
					//це файл без тіла, просто створити його та й все
					using (var fs = File.Create(destFullFileName))
					{
					}
				}
			}

			return destFullFileName;
		}


		static void _DownloadAttaches(IObjectDataManager obj, OID msgID, string destFolderPath)
		{
			//Отримаємо ідентифікатори файлів-вкладень, якщо вони є у листі:
			OID[] attachesIDs = obj.GetChildren(msgID, Foss.FossDoc.ApplicationServer.Messaging.Schema.Message.PR_MESSAGE_ATTACHMENTS);
			if (attachesIDs != null)
			{
				foreach (var attID in attachesIDs)
				{
					_DownloadAttach(obj, attID, destFolderPath, 0);

					//В інтерфейсній збірці бізнес-логіки є корисний утілітний клас для завантаження файлів-вкладень
					//Він робить все, щоб "завантажити" файл з сервера - ви також можете його використовувати
					//Foss.FossDoc.ExternalModules.BusinessLogic.Utils.Attach.DownloadAttach(session, attID, destFolderPath, null, 0);
				}
			}
		}

		static void _Log(string message, params object[] args)
		{
			Console.WriteLine(message, args);
		}

	}
}
