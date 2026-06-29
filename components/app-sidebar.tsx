"use client"

import * as React from "react"
import Link from "next/link"
import { usePathname } from "next/navigation"
import { 
  Terminal, 
  Activity, 
  Layers, 
  ArrowUpRight, 
  Sliders, 
  X
} from "lucide-react"

import {
  Sidebar,
  SidebarContent,
  SidebarGroup,
  SidebarGroupContent,
  SidebarHeader,
  SidebarMenu,
  SidebarMenuButton,
  SidebarMenuItem,
  useSidebar
} from "@/components/ui/sidebar"

const NAVIGATION_LINKS = [
  { title: "COMMAND_CENTER", url: "/command-center", icon: Terminal },
  { title: "OPERATIONS_MATRIX", url: "/operations", icon: Activity },
  { title: "AGENT_NETWORK", url: "/agent-network", icon: Layers },
  { title: "LOG_INGESTION", url: "/admin/field-reports", icon: ArrowUpRight },
  { title: "SYSTEM_CONFIG", url: "/systems", icon: Sliders },
]

export function AppSidebar() {
  const pathname = usePathname()
  const { toggleSidebar, open } = useSidebar()

  return (
    <Sidebar variant="sidebar" collapsible="icon" className="border-r border-neutral-900 bg-neutral-950">
      {/* Sidebar Header */}
      <SidebarHeader className="p-4 border-b border-neutral-900 flex flex-row items-center justify-between">
        <Link href="/" className="flex items-center gap-2 font-mono tracking-tighter">
          <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
          {open && <span className="font-black text-sm text-white tracking-widest">NEXUS//PRIME</span>}
        </Link>
        
        {/* Toggle / Closing Hint */}
        {open && (
          <button 
            onClick={toggleSidebar} 
            className="p-1 hover:bg-neutral-900 text-neutral-500 hover:text-white rounded transition-colors"
            title="Collapse Sidebar [Ctrl+B]"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        )}
      </SidebarHeader>
      
      {/* Sidebar Navigation Options */}
      <SidebarContent className="pt-2">
        <SidebarGroup>
          <SidebarGroupContent>
            <SidebarMenu>
              {NAVIGATION_LINKS.map((item) => {
                const isActive = pathname === item.url
                return (
                  <SidebarMenuItem key={item.title}>
                    <SidebarMenuButton asChild isActive={isActive} tooltip={item.title}>
                      <Link 
                        href={item.url} 
                        className={`flex items-center gap-3 px-3 py-2.5 font-mono text-[11px] uppercase tracking-wider transition-all duration-150 ${
                          isActive 
                            ? "bg-neutral-900 text-white border-l-2 border-emerald-500 font-bold" 
                            : "text-neutral-500 hover:bg-neutral-900/50 hover:text-neutral-200"
                        }`}
                      >
                        <item.icon className={`w-4 h-4 shrink-0 ${isActive ? "text-emerald-400" : "text-neutral-600"}`} />
                        <span>{item.title}</span>
                      </Link>
                    </SidebarMenuButton>
                  </SidebarMenuItem>
                )
              })}
            </SidebarMenu>
          </SidebarGroupContent>
        </SidebarGroup>
      </SidebarContent>
    </Sidebar>
  )
}