using System.IdentityModel.Tokens.Jwt;
using System.Security.Cryptography.X509Certificates;
using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace ExchangeEventGenerator; 

public class AzureClientConnection {
    
    //Azure/Graph API permissions:
    private static readonly string[] RequiredScopes =
    {
        "User.Read.All", "Calendars.ReadWrite"//, "Mail.ReadWrite"
    };
    
    public GraphServiceClient Client { get; }
    
    public AzureClientConnection(IConfiguration configuration) {
        Client = GetClient(configuration).Result;
    }
    
    /*
     * Establishes a connection to azure, authenticates the program, verifies permissions
     * Returns a static instance of a GraphServiceClient
     */
    private static async Task<GraphServiceClient> GetClient(IConfiguration configuration) {
        var credentials = configuration.GetSection("Credentials");
        var tenantId = credentials.GetValue<string>("TenantId");
        var clientId = credentials.GetValue<string>("ClientId");
        var certThumbprint = credentials.GetValue<string>("CertificateThumbprint");
        
        if(string.IsNullOrWhiteSpace(tenantId) || string.IsNullOrWhiteSpace(clientId) || 
           string.IsNullOrWhiteSpace(certThumbprint))
            throw new NullReferenceException("Malformed appsettings.json file.");

        //.default scope infers permissions directly from the application registration in azure
        var scopes = new[] { "https://graph.microsoft.com/.default" };
        
        var certificate = GetCertByThumbprint(StoreLocation.CurrentUser, certThumbprint);
        if(certificate == null)
            throw new NullReferenceException($"Could not find certificate with thumbprint: {certThumbprint}");
        var clientCertificateCredential = new ClientCertificateCredential(tenantId, clientId, certificate);
        
        //TODO add retry functionality
        Console.Write("Waiting for authentication...");
        try{
            //acquired tokens are cached so they get re-used during the GraphServiceClient object creation
            var token = await clientCertificateCredential.GetTokenAsync(new TokenRequestContext(scopes));
            VerifyGrantedPermissions(token);
        }
        catch (Exception e) when (e is AuthenticationFailedException or MsalServiceException){
            Console.WriteLine(e.Message);
            if(e is MsalServiceException)
                Console.WriteLine(e.StackTrace);
            Environment.Exit(-1);
        }
        Console.WriteLine("Success");
        return new GraphServiceClient(clientCertificateCredential, scopes);
    }

    //Decodes the given AccessToken and makes sure the azure permissions include all of the required permissions
    private static void VerifyGrantedPermissions(AccessToken jwtToken) {
        var handler = new JwtSecurityTokenHandler();
        var jsonToken = handler.ReadJwtToken(jwtToken.Token);
        var perms = jsonToken?.Claims?
            .Where(claim => claim.Type == "roles").ToArray()
            .Select(roles => roles.Value).ToArray();
        if(perms == null){
            throw new AuthenticationFailedException("Failed to acquire access token.");
        }
        foreach (var permission in RequiredScopes){
            if(!perms.Contains(permission)){
                throw new AuthenticationFailedException($"Failed to acquire required azure permission: {permission}");
            }
        }
    }
    
    /*
     * Searches the given StoreLocation for a certificate with the given thumbprint
     */
    private static X509Certificate2? GetCertByThumbprint(StoreLocation storeLoc, string thumbprint){
        var store = new X509Store(storeLoc);
        X509Certificate2? cer = null;
 
        store.Open(OpenFlags.ReadOnly);
 
        // look for the specific cert by thumbprint -- Also, for self signed, set the valid parameter to 'false'
        var certs = store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint, false);
        if (certs.Count > 0){
            cer = certs[0];
        }
        store.Close();

        return cer;
    }
}