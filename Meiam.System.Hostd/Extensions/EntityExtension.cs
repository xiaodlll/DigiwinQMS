using Meiam.System.Common.Helpers;
using Meiam.System.Model;
using System;

namespace Meiam.System.Hostd.Extensions
{
    public static class EntityExtension
    {
        public static TSource ToCreate<TSource>(this TSource source, UserSessionVM userSession)
        {
            var types = source.GetType();

            if (types.GetProperty("ID") != null)
            {
                types.GetProperty("ID").SetValue(source, SequentialGuid.Generate(), null);
            }

            if (types.GetProperty("CreateTime") != null)
            {
                types.GetProperty("CreateTime").SetValue(source, DateTime.Now, null);
            }

            if (types.GetProperty("UpdateTime") != null)
            {
                types.GetProperty("UpdateTime").SetValue(source, DateTime.Now, null);
            }

            if (types.GetProperty("CreateID") != null)
            {
                types.GetProperty("CreateID").SetValue(source, userSession.UserID, null);
            }

            if (types.GetProperty("UpdateID") != null)
            {
                types.GetProperty("UpdateID").SetValue(source, userSession.UserID, null);
            }


            return source;
        }

        public static TSource ToUpdate<TSource>(this TSource source, UserSessionVM userSession)
        {
            var types = source.GetType();

            if (types.GetProperty("UpdateTime") != null)
            {
                types.GetProperty("UpdateTime").SetValue(source, DateTime.Now, null);
            }

            if (types.GetProperty("UpdateID") != null)
            {
                types.GetProperty("UpdateID").SetValue(source, userSession.UserID, null);
            }

            return source;
        }

    }
}
