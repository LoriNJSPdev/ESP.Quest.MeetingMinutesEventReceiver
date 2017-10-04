
using System;
using System.Data.Objects;
using System.Web;
public static class ContextHelper<T> where T : ObjectContext, new()
{   
    
    #region Methods
    public static T GetCurrentContext(string ObjectContextKey) 
    { 
        HttpContext httpContext = HttpContext.Current;
        if (httpContext != null)
        { 
            string contextTypeKey = ObjectContextKey + typeof(T).Name;
            if (httpContext.Items[contextTypeKey] == null) 
            { 
                httpContext.Items.Add(contextTypeKey, new T());
            } 
            return httpContext.Items[contextTypeKey] as T; 
        }    
        throw new ApplicationException("There is no Http Context available");
    }  
    #endregion
}
