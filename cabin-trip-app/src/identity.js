import { createContext, useContext, useState, useCallback } from 'react';

export const IdentityContext = createContext(null);

export function useIdentity() {
  return useContext(IdentityContext);
}

export function useIdentityState() {
  const [personId, setPersonIdState] = useState(localStorage.getItem('cabin_person_id') || '');
  const [personName, setPersonName] = useState(localStorage.getItem('cabin_person_name') || '');

  const setPerson = useCallback((id, name) => {
    localStorage.setItem('cabin_person_id', id);
    localStorage.setItem('cabin_person_name', name);
    setPersonIdState(id);
    setPersonName(name);
  }, []);

  const clearPerson = useCallback(() => {
    localStorage.removeItem('cabin_person_id');
    localStorage.removeItem('cabin_person_name');
    setPersonIdState('');
    setPersonName('');
  }, []);

  return { personId, personName, setPerson, clearPerson };
}
