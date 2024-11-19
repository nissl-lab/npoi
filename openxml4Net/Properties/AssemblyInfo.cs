using System.Security;

#if NETSTANDARD2_1 || NET6_0_OR_GREATER || NETSTANDARD2_0 || NET40
[assembly: AllowPartiallyTrustedCallers]
#endif
#if NETSTANDARD2_1 || NET6_0_OR_GREATER || NETSTANDARD2_0 || NET40 || NET45
[assembly: SecurityRules(SecurityRuleSet.Level1)]
#endif
