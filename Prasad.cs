public static class ExceptionExtensions
    {
		public static string FormatException(this Exception ex)
		{
			if (ex is DbEntityValidationException)
			{
				return (ex as DbEntityValidationException).DbEntityValidationExceptionToString();
			}
			else if (ex is SocketException)
			{
				return $"Failed to establish connection: {ex.ExceptionToString()}";
			}

			return ex.ExceptionToString();
		}

		public static string ExceptionToString(this Exception ex)
		{
			StringBuilder sb = new StringBuilder(Environment.NewLine);

			sb.AppendLine($"Date/Time: {DateTime.Now.ToString()}");
			sb.AppendLine($"Exception Type: {ex.GetType().FullName}");
			sb.AppendLine($"Message: {ex.Message}");
			sb.AppendLine($"Source: {ex.Source}");

			// e.Data.Add("user", Thread.CurrentPrincipal.Identity.Name);
			foreach (var key in ex.Data.Keys)
			{
				sb.AppendLine($"{key.ToString()}: {ex.Data[key].ToString()}");
			}

			if (!ex.StackTrace.IsNullOrWhiteSpace())
			{
				sb.AppendLine($"Stack Trace: {ex.StackTrace}");
			}

			if (ex.InnerException != null)
			{
				sb.AppendLine($"Inner Exception: {ex.InnerException.ExceptionToString()}");
			}

			return sb.ToString();
		}

		/// <summary>
		/// A DbEntityValidationException extension method that formates validation errors to string.
		/// </summary>
		public static string DbEntityValidationExceptionToString(this DbEntityValidationException e)
		{
			var validationErrors = e.DbEntityValidationResultToString();
			var exceptionMessage = string.Format("{0}Validation errors:{1}{0}{2}", Environment.NewLine, validationErrors, e.ExceptionToString());
			return exceptionMessage;
		}

		/// <summary>
		/// A DbEntityValidationException extension method that aggregate database entity validation results to string.
		/// </summary>
		public static string DbEntityValidationResultToString(this DbEntityValidationException e)
		{
			return e.EntityValidationErrors
					.Select(dbEntityValidationResult => dbEntityValidationResult.DbValidationErrorsToString(dbEntityValidationResult.ValidationErrors))
					.Aggregate(string.Empty, (current, next) => string.Format("{0}{1}{2}", current, Environment.NewLine, next));
		}

		/// <summary>
		/// A DbEntityValidationResult extension method that to strings database validation errors.
		/// </summary>
		public static string DbValidationErrorsToString(this DbEntityValidationResult dbEntityValidationResult, IEnumerable<DbValidationError> dbValidationErrors)
		{
			var entityName = string.Format("[{0}]", dbEntityValidationResult.Entry.Entity.GetType().Name);
			const string indentation = "\t - ";
			var aggregatedValidationErrorMessages = dbValidationErrors.Select(error => string.Format("[{0} - {1}]", error.PropertyName, error.ErrorMessage))
												   .Aggregate(string.Empty, (current, validationErrorMessage) => current + (Environment.NewLine + indentation + validationErrorMessage));
			return string.Format("{0}{1}", entityName, aggregatedValidationErrorMessages);
		}
	}


===================
        public static bool AnyOrDefault<T>(this IEnumerable<T> list)
        {
            return list?.Any() ?? false;
        }

        public static bool AnyOrDefault<T>(this IEnumerable<T> list, Func<T, bool> match)
        {
            return list?.Any(match) ?? false;
        }
        
===================
public static class JsonConvertExtensions
    {
        public static string JsonSerialize<T>(this T jsonObj, JsonSerializerSettings settings = null)
        {
            return settings == null ? JsonConvert.SerializeObject(jsonObj) : JsonConvert.SerializeObject(jsonObj, settings);
        }

        public static T JsonDeserialize<T>(this string json)
        {
            return JsonConvert.DeserializeObject<T>(json);
        }

        public static T JsonDeserializePath<T>(this string filePath)
        {
            var json = File.ReadAllText(filePath);
            return JsonDeserialize<T>(json);
        }

        public static bool IsNullOrEmpty(this JToken token)
        {
            return (token == null) ||
                   (token.Type == JTokenType.Array && !token.HasValues) ||
                   (token.Type == JTokenType.Object && !token.HasValues) ||
                   (token.Type == JTokenType.String && token.ToString() == string.Empty) ||
                   (token.Type == JTokenType.Null);
        }

        public static bool JsonStringIsNullOrEmpty(this string token) => JToken.Parse(token).IsNullOrEmpty();

        public static T Clone<T>(this T source)
        {
            var serialized = JsonConvert.SerializeObject(source);
            return JsonConvert.DeserializeObject<T>(serialized);
        }
    }
=====================
        public static bool IsNullOrEmpty(this string str) => string.IsNullOrEmpty(str);

        public static bool IsNullOrWhiteSpace(this string str) => string.IsNullOrWhiteSpace(str);
=====================
        public static int GetCustomerId(this IIdentity identity)
        {
            var user = (identity as ClaimsIdentity);
            return user.GetClaim<int>(MSIClaimTypes.CustomerId);
        }
        
        public static T GetClaim<T>(this ClaimsIdentity identity, string claimType)
        {
            if (!identity.HasClaim(t => t.Type == claimType))
            {
                throw new HttpResponseException(HttpStatusCode.BadRequest);
            }
            var claim = identity.Claims.FirstOrDefault(c => c.Type == claimType);
            return (T)TypeDescriptor.GetConverter(typeof(T)).ConvertFromInvariantString(claim.Value);
        }        
        
