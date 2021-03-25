using System.Collections.Generic;
using EindOpdrachtLeandro.model;

namespace EindOpdrachtLeandro.Repositories
{
    interface Iadresgegeven
    {
        IList<Adresgegeven> GetAll();
    }
}
