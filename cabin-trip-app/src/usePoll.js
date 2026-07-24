import { useEffect, useState, useCallback, useRef } from 'react';
import { api } from './api';

// Polls a GET endpoint every `interval` ms and exposes data + a manual refresh + local setter for optimistic updates.
export function usePoll(url, interval = 10000) {
  const [data, setData] = useState(null);
  const [error, setError] = useState(null);
  const timer = useRef(null);

  const load = useCallback(async () => {
    try {
      const d = await api.get(url);
      setData(d);
      setError(null);
    } catch (e) {
      setError(e.message);
    }
  }, [url]);

  useEffect(() => {
    load();
    timer.current = setInterval(load, interval);
    return () => clearInterval(timer.current);
  }, [load, interval]);

  return { data, setData, error, reload: load };
}
