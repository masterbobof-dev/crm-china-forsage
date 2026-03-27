import { useState, useEffect, useCallback, useRef } from 'react';
import { supabase } from '../supabase';

export const toSnakeCase = (obj: any): any => {
  if (Array.isArray(obj)) {
    return obj.map(v => toSnakeCase(v));
  } else if (obj !== null && obj.constructor === Object) {
    return Object.keys(obj).reduce((result, key) => {
      // Exclude UI-only properties that are not in the database schema
      if (key === 'isWaybillHeader') {
        return result;
      }
      const snakeKey = key.replace(/([A-Z])/g, "_$1").toLowerCase();
      result[snakeKey] = toSnakeCase(obj[key]);
      return result;
    }, {} as any);
  }
  return obj;
};

export const toCamelCase = (obj: any): any => {
  if (Array.isArray(obj)) {
    return obj.map(v => toCamelCase(v));
  } else if (obj !== null && obj.constructor === Object) {
    return Object.keys(obj).reduce((result, key) => {
      const camelKey = key.replace(/_([a-z])/g, (g) => g[1].toUpperCase());
      result[camelKey] = toCamelCase(obj[key]);
      return result;
    }, {} as any);
  }
  return obj;
};

export function useSupabaseSync<T extends { id: string }>(tableName: string, initialData: T[]) {
  const [data, setData] = useState<T[]>(initialData);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    fetchData();
  }, []);

  const fetchData = async () => {
    try {
      const { data: fetchedData, error } = await supabase.from(tableName).select('*');
      if (error) throw error;
      if (fetchedData && fetchedData.length > 0) {
        setData(toCamelCase(fetchedData));
      }
    } catch (error) {
      console.error(`Error fetching ${tableName}:`, error);
    } finally {
      setLoading(false);
    }
  };

  const syncData = useCallback((newDataOrUpdater: T[] | ((prev: T[]) => T[])) => {
    setData(prev => {
      const newData = typeof newDataOrUpdater === 'function' ? (newDataOrUpdater as any)(prev) : newDataOrUpdater;
      
      const prevIds = new Set(prev.map(item => item.id));
      const newIds = new Set(newData.map(item => item.id));
      const deletedIds = [...prevIds].filter(id => !newIds.has(id));

      Promise.resolve().then(async () => {
        try {
          if (deletedIds.length > 0) {
            const { error: deleteError } = await supabase.from(tableName).delete().in('id', deletedIds);
            if (deleteError) throw deleteError;
          }
          if (newData.length > 0) {
            const snakeCaseData = toSnakeCase(newData);
            const { error: upsertError } = await supabase.from(tableName).upsert(snakeCaseData);
            if (upsertError) throw upsertError;
          }
        } catch (error) {
          console.error(`Error syncing ${tableName}:`, error);
        }
      });
      
      return newData;
    });
  }, [tableName]);

  return [data, syncData, loading] as const;
}
