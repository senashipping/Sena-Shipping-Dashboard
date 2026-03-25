import { useState, useEffect } from 'react';

/**
 * Custom hook for debouncing values
 * @param value - The value to debounce
 * @param delay - Delay in milliseconds
 * @returns Debounced value
 */
export function useDebounce<T>(value: T, delay: number): T {
  const [debouncedValue, setDebouncedValue] = useState<T>(value);

  useEffect(() => {
    const handler = setTimeout(() => {
      setDebouncedValue(value);
    }, delay);

    return () => {
      clearTimeout(handler);
    };
  }, [value, delay]);

  return debouncedValue;
}

/**
 * Custom hook for client-side search with debouncing
 * @param data - Array of data to search through
 * @param searchTerm - Current search term
 * @param searchFields - Array of field paths to search in (e.g., ['name', 'email', 'ship.name'])
 * @param debounceMs - Debounce delay in milliseconds (default: 300)
 * @returns Filtered data based on search term
 */
export function useClientSearch<T extends Record<string, any>>(
  data: T[],
  searchTerm: string,
  searchFields: string[],
  debounceMs: number = 300
): T[] {
  const debouncedSearchTerm = useDebounce(searchTerm, debounceMs);

  return data.filter((item) => {
    if (!debouncedSearchTerm.trim()) return true;

    const searchLower = debouncedSearchTerm.toLowerCase();
    
    return searchFields.some((fieldPath) => {
      const value = getNestedValue(item, fieldPath);
      return value && String(value).toLowerCase().includes(searchLower);
    });
  });
}

/**
 * Helper function to get nested object values using dot notation
 * @param obj - Object to search in
 * @param path - Dot notation path (e.g., 'user.name', 'ship.imoNumber')
 * @returns Value at the specified path
 */
function getNestedValue(obj: any, path: string): any {
  return path.split('.').reduce((current, key) => {
    return current && current[key] !== undefined ? current[key] : null;
  }, obj);
}