export type Json =
  | string
  | number
  | boolean
  | null
  | { [key: string]: Json | undefined }
  | Json[]

export type Database = {
  // Allows to automatically instantiate createClient with right options
  // instead of createClient<Database, { PostgrestVersion: 'XX' }>(URL, KEY)
  __InternalSupabase: {
    PostgrestVersion: "14.5"
  }
  public: {
    Tables: {
      excel_job_files: {
        Row: {
          created_at: string
          id: string
          job_id: string
          original_name: string
          role: Database["public"]["Enums"]["file_role"]
          sheet_meta: Json | null
          size_bytes: number
          storage_path: string
          user_id: string
        }
        Insert: {
          created_at?: string
          id?: string
          job_id: string
          original_name: string
          role?: Database["public"]["Enums"]["file_role"]
          sheet_meta?: Json | null
          size_bytes?: number
          storage_path: string
          user_id: string
        }
        Update: {
          created_at?: string
          id?: string
          job_id?: string
          original_name?: string
          role?: Database["public"]["Enums"]["file_role"]
          sheet_meta?: Json | null
          size_bytes?: number
          storage_path?: string
          user_id?: string
        }
        Relationships: [
          {
            foreignKeyName: "excel_job_files_job_id_fkey"
            columns: ["job_id"]
            isOneToOne: false
            referencedRelation: "excel_jobs"
            referencedColumns: ["id"]
          },
        ]
      }
      excel_jobs: {
        Row: {
          ai_plan: Json | null
          completed_at: string | null
          created_at: string
          error: string | null
          id: string
          intent: string | null
          kind: Database["public"]["Enums"]["job_kind"]
          name: string
          output_name: string | null
          output_path: string | null
          params: Json
          stats: Json
          status: Database["public"]["Enums"]["job_status"]
          updated_at: string
          user_id: string
          warnings: Json
        }
        Insert: {
          ai_plan?: Json | null
          completed_at?: string | null
          created_at?: string
          error?: string | null
          id?: string
          intent?: string | null
          kind?: Database["public"]["Enums"]["job_kind"]
          name?: string
          output_name?: string | null
          output_path?: string | null
          params?: Json
          stats?: Json
          status?: Database["public"]["Enums"]["job_status"]
          updated_at?: string
          user_id: string
          warnings?: Json
        }
        Update: {
          ai_plan?: Json | null
          completed_at?: string | null
          created_at?: string
          error?: string | null
          id?: string
          intent?: string | null
          kind?: Database["public"]["Enums"]["job_kind"]
          name?: string
          output_name?: string | null
          output_path?: string | null
          params?: Json
          stats?: Json
          status?: Database["public"]["Enums"]["job_status"]
          updated_at?: string
          user_id?: string
          warnings?: Json
        }
        Relationships: []
      }
      user_roles: {
        Row: {
          created_at: string
          id: string
          role: Database["public"]["Enums"]["app_role"]
          user_id: string
        }
        Insert: {
          created_at?: string
          id?: string
          role?: Database["public"]["Enums"]["app_role"]
          user_id: string
        }
        Update: {
          created_at?: string
          id?: string
          role?: Database["public"]["Enums"]["app_role"]
          user_id?: string
        }
        Relationships: []
      }
      workspace_files: {
        Row: {
          created_at: string
          id: string
          inspector: Json | null
          original_name: string
          sheet_meta: Json | null
          size_bytes: number
          storage_path: string
          user_id: string
          workspace_id: string
        }
        Insert: {
          created_at?: string
          id?: string
          inspector?: Json | null
          original_name: string
          sheet_meta?: Json | null
          size_bytes?: number
          storage_path: string
          user_id: string
          workspace_id: string
        }
        Update: {
          created_at?: string
          id?: string
          inspector?: Json | null
          original_name?: string
          sheet_meta?: Json | null
          size_bytes?: number
          storage_path?: string
          user_id?: string
          workspace_id?: string
        }
        Relationships: [
          {
            foreignKeyName: "workspace_files_workspace_id_fkey"
            columns: ["workspace_id"]
            isOneToOne: false
            referencedRelation: "workspaces"
            referencedColumns: ["id"]
          },
        ]
      }
      workspace_messages: {
        Row: {
          content: string
          created_at: string
          id: string
          role: string
          tool_data: Json | null
          user_id: string
          workspace_id: string
        }
        Insert: {
          content?: string
          created_at?: string
          id?: string
          role: string
          tool_data?: Json | null
          user_id: string
          workspace_id: string
        }
        Update: {
          content?: string
          created_at?: string
          id?: string
          role?: string
          tool_data?: Json | null
          user_id?: string
          workspace_id?: string
        }
        Relationships: [
          {
            foreignKeyName: "workspace_messages_workspace_id_fkey"
            columns: ["workspace_id"]
            isOneToOne: false
            referencedRelation: "workspaces"
            referencedColumns: ["id"]
          },
        ]
      }
      workspace_versions: {
        Row: {
          created_at: string
          id: string
          label: string
          output_name: string | null
          output_path: string | null
          parent_version_id: string | null
          plan: Json | null
          size_bytes: number
          stats: Json
          user_id: string
          version_number: number
          warnings: Json
          workspace_id: string
        }
        Insert: {
          created_at?: string
          id?: string
          label?: string
          output_name?: string | null
          output_path?: string | null
          parent_version_id?: string | null
          plan?: Json | null
          size_bytes?: number
          stats?: Json
          user_id: string
          version_number: number
          warnings?: Json
          workspace_id: string
        }
        Update: {
          created_at?: string
          id?: string
          label?: string
          output_name?: string | null
          output_path?: string | null
          parent_version_id?: string | null
          plan?: Json | null
          size_bytes?: number
          stats?: Json
          user_id?: string
          version_number?: number
          warnings?: Json
          workspace_id?: string
        }
        Relationships: [
          {
            foreignKeyName: "workspace_versions_parent_version_id_fkey"
            columns: ["parent_version_id"]
            isOneToOne: false
            referencedRelation: "workspace_versions"
            referencedColumns: ["id"]
          },
          {
            foreignKeyName: "workspace_versions_workspace_id_fkey"
            columns: ["workspace_id"]
            isOneToOne: false
            referencedRelation: "workspaces"
            referencedColumns: ["id"]
          },
        ]
      }
      workspaces: {
        Row: {
          created_at: string
          id: string
          name: string
          updated_at: string
          user_id: string
        }
        Insert: {
          created_at?: string
          id?: string
          name?: string
          updated_at?: string
          user_id: string
        }
        Update: {
          created_at?: string
          id?: string
          name?: string
          updated_at?: string
          user_id?: string
        }
        Relationships: []
      }
    }
    Views: {
      [_ in never]: never
    }
    Functions: {
      has_role: {
        Args: {
          _role: Database["public"]["Enums"]["app_role"]
          _user_id: string
        }
        Returns: boolean
      }
    }
    Enums: {
      app_role: "admin" | "user"
      file_role: "input" | "output"
      job_kind: "merge" | "dedupe" | "diff" | "format" | "summary" | "auto"
      job_status:
        | "draft"
        | "queued"
        | "planning"
        | "running"
        | "succeeded"
        | "failed"
    }
    CompositeTypes: {
      [_ in never]: never
    }
  }
}

type DatabaseWithoutInternals = Omit<Database, "__InternalSupabase">

type DefaultSchema = DatabaseWithoutInternals[Extract<keyof Database, "public">]

export type Tables<
  DefaultSchemaTableNameOrOptions extends
    | keyof (DefaultSchema["Tables"] & DefaultSchema["Views"])
    | { schema: keyof DatabaseWithoutInternals },
  TableName extends DefaultSchemaTableNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof (DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"] &
        DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Views"])
    : never = never,
> = DefaultSchemaTableNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? (DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"] &
      DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Views"])[TableName] extends {
      Row: infer R
    }
    ? R
    : never
  : DefaultSchemaTableNameOrOptions extends keyof (DefaultSchema["Tables"] &
        DefaultSchema["Views"])
    ? (DefaultSchema["Tables"] &
        DefaultSchema["Views"])[DefaultSchemaTableNameOrOptions] extends {
        Row: infer R
      }
      ? R
      : never
    : never

export type TablesInsert<
  DefaultSchemaTableNameOrOptions extends
    | keyof DefaultSchema["Tables"]
    | { schema: keyof DatabaseWithoutInternals },
  TableName extends DefaultSchemaTableNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"]
    : never = never,
> = DefaultSchemaTableNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"][TableName] extends {
      Insert: infer I
    }
    ? I
    : never
  : DefaultSchemaTableNameOrOptions extends keyof DefaultSchema["Tables"]
    ? DefaultSchema["Tables"][DefaultSchemaTableNameOrOptions] extends {
        Insert: infer I
      }
      ? I
      : never
    : never

export type TablesUpdate<
  DefaultSchemaTableNameOrOptions extends
    | keyof DefaultSchema["Tables"]
    | { schema: keyof DatabaseWithoutInternals },
  TableName extends DefaultSchemaTableNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"]
    : never = never,
> = DefaultSchemaTableNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"][TableName] extends {
      Update: infer U
    }
    ? U
    : never
  : DefaultSchemaTableNameOrOptions extends keyof DefaultSchema["Tables"]
    ? DefaultSchema["Tables"][DefaultSchemaTableNameOrOptions] extends {
        Update: infer U
      }
      ? U
      : never
    : never

export type Enums<
  DefaultSchemaEnumNameOrOptions extends
    | keyof DefaultSchema["Enums"]
    | { schema: keyof DatabaseWithoutInternals },
  EnumName extends DefaultSchemaEnumNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof DatabaseWithoutInternals[DefaultSchemaEnumNameOrOptions["schema"]]["Enums"]
    : never = never,
> = DefaultSchemaEnumNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? DatabaseWithoutInternals[DefaultSchemaEnumNameOrOptions["schema"]]["Enums"][EnumName]
  : DefaultSchemaEnumNameOrOptions extends keyof DefaultSchema["Enums"]
    ? DefaultSchema["Enums"][DefaultSchemaEnumNameOrOptions]
    : never

export type CompositeTypes<
  PublicCompositeTypeNameOrOptions extends
    | keyof DefaultSchema["CompositeTypes"]
    | { schema: keyof DatabaseWithoutInternals },
  CompositeTypeName extends PublicCompositeTypeNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof DatabaseWithoutInternals[PublicCompositeTypeNameOrOptions["schema"]]["CompositeTypes"]
    : never = never,
> = PublicCompositeTypeNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? DatabaseWithoutInternals[PublicCompositeTypeNameOrOptions["schema"]]["CompositeTypes"][CompositeTypeName]
  : PublicCompositeTypeNameOrOptions extends keyof DefaultSchema["CompositeTypes"]
    ? DefaultSchema["CompositeTypes"][PublicCompositeTypeNameOrOptions]
    : never

export const Constants = {
  public: {
    Enums: {
      app_role: ["admin", "user"],
      file_role: ["input", "output"],
      job_kind: ["merge", "dedupe", "diff", "format", "summary", "auto"],
      job_status: [
        "draft",
        "queued",
        "planning",
        "running",
        "succeeded",
        "failed",
      ],
    },
  },
} as const
