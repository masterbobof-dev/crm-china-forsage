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
      console.log(`Fetching data from ${tableName}...`);
      const { data: fetchedData, error } = await supabase.from(tableName).select('*');
      if (error) {
        console.error(`Supabase fetch error for ${tableName}:`, error);
        throw error;
      }
      console.log(`Fetched ${fetchedData?.length || 0} items from ${tableName}`);
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
    console.log(`[SupabaseSync] syncData called for ${tableName}`);
    setData(prev => {
      const newData = typeof newDataOrUpdater === 'function' ? (newDataOrUpdater as any)(prev) : newDataOrUpdater;
      
      const prevIds = new Set(prev.map(item => item.id));
      const newIds = new Set(newData.map(item => item.id));
      const deletedIds = [...prevIds].filter(id => !newIds.has(id));

      Promise.resolve().then(async () => {
        try {
          if (deletedIds.length > 0) {
            console.log(`Deleting ${deletedIds.length} items from ${tableName}...`);
            const { error: deleteError } = await supabase.from(tableName).delete().in('id', deletedIds);
            if (deleteError) {
              console.error(`Supabase delete error for ${tableName}:`, deleteError);
              throw deleteError;
            }
          }
          if (newData.length > 0) {
            console.log(`Upserting ${newData.length} items to ${tableName}...`);
            try {
              const snakeCaseData = toSnakeCase(newData);
              console.log(`[SupabaseSync] snakeCaseData generated for ${tableName}`);
              const { error: upsertError } = await supabase.from(tableName).upsert(snakeCaseData);
              if (upsertError) {
                console.error(`Supabase upsert error for ${tableName}:`, upsertError);
                throw upsertError;
              }
              console.log(`Successfully synced ${tableName}`);
            } catch (e) {
              console.error(`[SupabaseSync] Error during upsert for ${tableName}:`, e);
              console.dir(e);
              throw e;
            }
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
