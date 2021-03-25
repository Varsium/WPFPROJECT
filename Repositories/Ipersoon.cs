using System.Collections.Generic;
using EindOpdrachtLeandro.model;

namespace EindOpdrachtLeandro.Repositories
{
    interface Ipersoon
    {
        IList<Persoon> GetAll();
    }
}
